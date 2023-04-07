#include "stdafx.h"
#include <fstream>
#include <codecvt>

#define _USE_MATH_DEFINES
#include <math.h>

#include "common.h"
#include "utils.h"

void Utils::SanitizeString( LPWSTR str, LPCWSTR notallow, WCHAR replace )
{
	LPWSTR pstr = str;
	while((pstr = wcspbrk(pstr, notallow)) != NULL) {
		*pstr = replace;
		pstr++;
	}
}

BOOL Utils::IsValidGUID( LPCWSTR guidStr, BOOL embraced )
{
	if (embraced && (*guidStr++ != '{'))
		return FALSE;

	const WCHAR GLT[] = L"00000000-0000-0000-0000-000000000000";
	for(SIZE_T pos = 0; pos < _countof(GLT) - 1; pos++) {
		WCHAR ch = *guidStr++;
		if (GLT[pos] == L'0') {
			if (!iswxdigit(ch))
				return FALSE;
		} else if (GLT[pos] != ch)
			return FALSE;
	}
	if (embraced && (*guidStr++ != '}'))
		return FALSE;
	return TRUE;
}

enum {
	PP_NONE		= 0x00,
	PP_PATH		= 0x01,
	PP_FILENS	= 0x02,
	PP_DEVNS	= 0x04,
	PP_UNCPR	= 0x08,
	PP_UNC		= 0x10,
	PP_DRIVE	= 0x20,
	PP_VOLUME	= 0x40,
	PP_PIPE		= 0x80,
	PP_SLASH	= 0x100
};
INT Utils::ParsePath( LPCWSTR path, SIZE_T maxlen, LPWSTR * endptr )
{
	typedef const struct {
		LPCWSTR seq;
		UINT len;
		UINT type;
		UINT next;
		BOOL ppart; // can be part of path
	} pathseq_t;
#define ADD_SEQ(seq) seq, _countof(seq) - 1/*\0*/
	pathseq_t sequences[] = {
			{ ADD_SEQ(L"\\\\?\\"), PP_FILENS, PP_DRIVE|PP_VOLUME|PP_UNC, FALSE },
			{ ADD_SEQ(L"\\\\.\\"), PP_DEVNS, PP_DRIVE|PP_PIPE|PP_PATH, FALSE }, // PP_PATH for device path
			{ ADD_SEQ(L"\\\\"), PP_UNCPR, PP_PATH, TRUE }, // must be after both PP_FILENS and PP_DEVNS
			{ ADD_SEQ(L"UNC\\"), PP_UNC, PP_PATH, TRUE },
			{ ADD_SEQ(L"x:"), PP_DRIVE, PP_PATH, FALSE },
			{ ADD_SEQ(L"Volume{00000000-0000-0000-0000-000000000000}\\"), PP_VOLUME, PP_PATH, TRUE },
			{ ADD_SEQ(L"pipe\\"), PP_PIPE, PP_PATH, TRUE },
			// must be the last
			{ NULL, PP_PATH, PP_PATH }
	};
	INT next = PP_PATH | PP_FILENS | PP_DEVNS | PP_UNCPR | PP_DRIVE;
	INT ppath = PP_NONE;
	LPCWSTR savepath = NULL;
	pathseq_t *i;
	while(maxlen && *path) {
		for(i = sequences; i->seq != NULL; i++) {
			if (maxlen < i->len)
				continue;
			if (i->type == PP_DRIVE) {
				if (!iswalpha(*path))
					continue;
				if (_wcsnicmp(i->seq+1, path+1, i->len-1) != 0)
					continue;
			} else if (i->type == PP_VOLUME) {
				if (_wcsnicmp(i->seq, path, 6) != 0)
					continue;
				if (!IsValidGUID(path + 6, TRUE))
					return -1;
				if (_wcsnicmp(i->seq + 44, path + 44, i->len - 44) != 0) // MSFT: All volume and mounted folder functions that take a volume GUID path as an input parameter require the trailing backslash
					return -1;
			} else {
				if (_wcsnicmp(i->seq, path, i->len) != 0)
					continue;
			}
			if ((i->type & next) == 0) {
				if ((next & PP_PATH) && i->ppart) {
					savepath = path;
					next = PP_PATH;
				} else
					return -1;
			} else
				ppath |= i->type;
			next = i->next;
			path += i->len;
			maxlen -= i->len;
			break;
		}
		if (i->seq == NULL) {
			if ((next & PP_PATH) == 0)
				return -1;
			if (savepath != NULL) path = savepath;
			if ((ppath & PP_DRIVE) && (*path == '\\')) { // consume only one separator, leaving others to normalizer
				ppath |= PP_SLASH;
				path++;
			}
			break;
		}
	}
	if (endptr != NULL)
		*endptr = (LPWSTR)path;

	return ppath;
}

BOOL Utils::IsPathRelative( INT ppath )
{
	switch (ppath) {
	case PP_NONE:
	case PP_SLASH:
		return TRUE;
	}
	return FALSE;
}

BOOL Utils::IsPathRelative( LPCWSTR path )
{
	INT ppath = ParsePath( path, wcslen(path) );
	if (ppath < 0)
		return TRUE;
	return IsPathRelative( ppath );
}

BOOL Utils::IsPathNetwork( INT ppath )
{
	switch (ppath) {
	case PP_UNCPR:
	case PP_FILENS|PP_UNC:
		return TRUE;
	}
	return FALSE;
}

BOOL Utils::IsPathNetwork( LPCWSTR path )
{
	INT ppath = ParsePath( path, wcslen(path) );
	if (ppath < 0)
		return FALSE;
	return IsPathNetwork( ppath );
}

BOOL Utils::IsPathDevice( INT ppath )
{
	switch (ppath) {
	case PP_DEVNS:
		return TRUE;
	}
	return FALSE;
}

BOOL Utils::IsPathDevice( LPCWSTR path )
{
	INT ppath = ParsePath( path, wcslen(path) );
	if (ppath < 0)
		return FALSE;
	return IsPathDevice( ppath );
}

BOOL Utils::IsPathPipe( INT ppath )
{
	switch (ppath) {
	case PP_DEVNS|PP_PIPE:
		return TRUE;
	}
	return FALSE;
}

BOOL Utils::IsPathPipe( LPCWSTR path )
{
	INT ppath = ParsePath( path, wcslen(path) );
	if (ppath < 0)
		return FALSE;
	return IsPathPipe( ppath );
}

BOOL Utils::SanitizeFileSystemString( LPWSTR str, SIZE_T strlen, WCHAR replacement, BOOL asfilename, BOOL notrailsep, LPWSTR *rootptr )
{
	LPWSTR in = str;
	INT ppath = PP_NONE;
	if (strlen == 0)
		strlen = wcslen(str);
	BOOL isnet = FALSE;
	BOOL isdev = FALSE;
	LPWSTR root = str;

	if (!asfilename) {
		ppath = ParsePath(str, strlen, &in);
		isnet = IsPathNetwork(ppath);
		isdev = IsPathDevice(ppath);						// burden logic
		if ((ppath < 0) ||									// invalid path
			((ppath & (PP_DRIVE|PP_SLASH)) == PP_DRIVE) ||	// don't support silly paths relative to current directory of the drive
			((ppath & (PP_VOLUME|PP_PIPE)) != 0) ||			// volume and pipe specifications are not supported
			isdev)											// only drive names in the device namespace are supported
			return FALSE;
		strlen -= (in - str);
		root = in;
	}

	INT tokens = 0;
	INT tokpos = 0;
	LPWSTR out = in;
	WCHAR ch;
	while(strlen && (ch = *in)) {
		if (ch < 32)
			ch = replacement;
		else {
			if ((tokpos == 0) && (ch == L' ')) {		// trim leading spaces
				in++;
				continue;
			}
			LPCWSTR forbidden;
			if (isnet && (tokens == 0)) {
				if ((tokpos == 0) && (ch == '.'))		// MSFT: (NetBIOS computer) Names can contain a period, but names cannot start with a period.
					return FALSE;
				forbidden = L"<>:\"\\/|?*";				// computer name
			} else
				forbidden = L"<>:\"\\/|?*";				// device/share/path
			if (wcschr(forbidden, ch)) {
				if (!asfilename && (ch == '\\')) {
					if ((in > str) && *(in - 1) == '\\') { // normalize separators
						in++;
					    continue;
					}
					while(out > root) {					// trim token trailing spaces and periods, preserving special . and .. tokens
						switch(*(out - 1)) {			
						case L'.':
							if ((tokpos == 1) ||
								(tokpos == 2) && (*(out - 2) == L'.'))
								break;
							out--;
							continue;
						case L' ':
							out--; tokpos--;
							continue;
						case L'\\':						// eliminate resuling empty token
							out--; tokens--;
							break;
						}
						break;
					}
					tokens++; tokpos = -1;
					if ((isdev && (tokens == 1)) ||		// device path
						(isnet && (tokens == 2)))		// or network path
						root = out + 1;					// update root
				} else {
					if (isnet && (tokens == 0))
						return FALSE;					// bad computer name
					ch = replacement;
				}
			}
		}
		*out++ = ch;
		in++;
		strlen--;
		tokpos++;
	}
	while(out > root) {									// trim trailing separators, spaces and periods
		switch(*(out - 1)) {
		case L'.':
			if ((tokpos == 1) ||
				(tokpos == 2) && (*(out - 2) == L'.'))
				break;
			out--;
			continue;
		case L' ':
			out--; tokpos--;
			continue;
		case L'\\':
			if (notrailsep ||							// trim trailing separator if requested
				((out - root) == 1) ||					// or final path is empty
				(*(in - 1) != '\\'))					// or initial path doesn't end with a trailing separator
				out--; tokens--;
			break;
		}
		break;
	}
	*out = '\0';

	if (rootptr)
		*rootptr = root;

	return TRUE;
}

BOOL Utils::SanitizeFileSystemString( CString & str, WCHAR replacement, BOOL filename, BOOL notrailsep, LPWSTR * rootptr )
{
	BOOL ret = SanitizeFileSystemString(str.LockBuffer(), str.GetLength(), replacement, filename, notrailsep, rootptr);
	str.UnlockBuffer();
	return ret;
}

LPWSTR Utils::AbsPath( LPCWSTR path, LPCWSTR reference, BOOL reflock, WCHAR sareplacement )
{
	if (IsPathRelative(reference))
		return NULL;

	SIZE_T pathlen = wcslen(path);
	BOOL relp = IsPathRelative(path);
	SIZE_T realsz = (relp?wcslen(reference):0) + pathlen + 2; // + separator + \0
	LPWSTR real = new WCHAR[realsz];
	if (real == NULL)
		return NULL;
	LPWSTR realroot = real;
	LPWSTR realptr = real;
	if (relp) {
		wcscpy_s(real, realsz, reference);
		if (!SanitizeFileSystemString(realptr, realsz - 1, sareplacement, FALSE, TRUE, &realroot)) {
			delete[] real;
			return NULL;
		}
		realptr += wcslen(real);
		if (pathlen > 0) {
			*realptr++ = '\\';
			*realptr = '\0';
		}
		if (reflock)
			realroot = realptr;
		if (*path == '\\') { // relative to root
			realptr = realroot;
			path++;
		}
	}

	SIZE_T pathsz = realsz - (realptr - real);
	wcscpy_s(realptr, pathsz, path);
	if (!SanitizeFileSystemString(realptr, pathsz - 1, sareplacement, FALSE, TRUE, !relp?&realroot:NULL)) {
		delete[] real;
		return NULL;
	}
	realptr = realroot;
	SIZE_T toklen = 0;
	WCHAR ch;
	for(LPWSTR pathptr = realroot; (ch = *pathptr) != '\0'; pathptr += toklen) {
		if (ch == L'\\')
			pathptr++;
		toklen = wcscspn(pathptr, L"\\");
		if (toklen == 0)
			continue;
		else if (*pathptr == L'.') {
			if (toklen == 1)
				continue;
			if ((toklen == 2) && (pathptr[1] == L'.')) { // parent dir
				INT count = 0;
				while(realptr > realroot) {
					realptr--;
					if (*realptr == L'\\' && (++count > 1)) {
						realptr++;
						break;
					}
				}
				continue;
			}
		}
		if (realptr != pathptr)
			wmemmove(realptr, pathptr, toklen + 1); // + separator or \0
		realptr += toklen + 1;
	}
	*realptr = '\0';

	return real;
}

/// (c)Cygon, http://blog.nuclex-games.com/2012/06/how-to-create-directories-recursively-with-win32/
/// <summary>Creates all directories down to the specified path</summary>
/// <param name="directory">Directory that will be created recursively</param>
/// <remarks>
///   The provided directory must not be terminated with a path separator.
/// </remarks>
BOOL Utils::CreateDirectoryRecursively( const CString directory )
{
	// If the specified directory name doesn't exist, do our thing
	DWORD fileAttributes = ::GetFileAttributesW((LPCWSTR)directory);
	if(fileAttributes == INVALID_FILE_ATTRIBUTES) {
		// Recursively do it all again for the parent directory, if any
		INT slashIndex = directory.ReverseFind('\\');
		if (slashIndex == -1)
			slashIndex = directory.ReverseFind('/');
		if (slashIndex != -1) {
			CreateDirectoryRecursively(directory.Left(slashIndex));
		}

		// Create the last directory on the path (the recursive calls will have taken
		// care of the parent directories by now)
		return CreateDirectoryW((LPCWSTR)directory, NULL);
	} else { // Specified directory name already exists as a file or directory
		bool isDirectoryOrJunction =
			((fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) ||
			((fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0);
 
		if(!isDirectoryOrJunction) {
			SetLastError(ERROR_CANNOT_MAKE);
			return FALSE;
		}
	}

	return TRUE;
}

DWORD Utils::GetModulePathName( LPWSTR pathname, SIZE_T pnsz, SIZE_T *nameidx /* = NULL */ )
{
	HMODULE hm;
	if (!GetModuleHandleExW( GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | 
							GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
							(LPWSTR) &GetModulePathName, &hm ) )
		return FALSE;

	DWORD len = GetModuleFileNameW( hm, pathname, pnsz );
	if ((len == 0) || (len == pnsz))
		return 0;

	if (nameidx) {
		LPWSTR mname = pathname;
		for(LPWSTR mptr = pathname; *mptr != '\0' && mptr < pathname + pnsz; mptr++)
			if (*mptr == '\\') mname = mptr + 1;
		*nameidx = mname - pathname;
	}

	return len;
}

BOOL Utils::Str2Clipboard(LPCWSTR str)
{
	BOOL open;
	if (open = OpenClipboard(NULL)) {
		SIZE_T len = (wcslen(str) + 1);
		HGLOBAL hMem =  GlobalAlloc(GMEM_MOVEABLE, len * sizeof(WCHAR));
		if (hMem == NULL) goto fail;
		wcscpy_s(static_cast<LPWSTR>(GlobalLock(hMem)), len, str);
		if (GlobalUnlock(hMem) || (GetLastError() != NO_ERROR)) goto fail;
		if (!EmptyClipboard()) goto fail;
		if (SetClipboardData(CF_UNICODETEXT, hMem) == NULL) goto fail;
		CloseClipboard();

		return TRUE;
	}

fail:
	if (open) {
		DWORD err = GetLastError();
		CloseClipboard();
		SetLastError(err);
	}
	return FALSE;
}

BOOL Utils::ParseDouble( LPCWSTR str, LPCWSTR & endptr, DOUBLE & val, _locale_t locale )
{
	WCHAR *eptr;
	DOUBLE v;
	errno = 0;
	if (locale != NULL)
		v = _wcstod_l(str, &eptr, locale);
	else
		v = wcstod(str, &eptr);
	if ((errno == ERANGE) ||
		(eptr == str))
		return FALSE;
	val = v;
	endptr = eptr;
	return TRUE;
}

double Utils::RadtoDeg(double Rad)
{
	return Rad / M_PI * 180.0;
}

double Utils::DegtoRad(double Deg)
{
    return Deg * M_PI / 180;
}

UINT Utils::GetFileName( BOOL save, LPCWSTR defpathname, LPCWSTR title, LPCWSTR suffix, LPCWSTR filter, LPCWSTR defext, OPENFILENAME* & filename )
{
	BOOL alloc = FALSE;

	if (filename == NULL) {
		filename = new OPENFILENAME;
		SecureZeroMemory(filename, sizeof(OPENFILENAME));
		alloc = TRUE;
	}
	if (filename == NULL)
		return LIBSTATUS_SYSERR_NO_MEM;

	filename->lStructSize = sizeof(OPENFILENAME);
	filename->hwndOwner = (HWND)GetHWindow();
	filename->lpstrFilter = filter;
	filename->lpstrTitle = title;
	if (save)
		filename->Flags = OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_NOCHANGEDIR|OFN_OVERWRITEPROMPT|OFN_PATHMUSTEXIST;
	else
		filename->Flags = OFN_DONTADDTORECENT|OFN_ENABLESIZING|OFN_NOCHANGEDIR|OFN_FILEMUSTEXIST|OFN_PATHMUSTEXIST|OFN_EXPLORER|OFN_HIDEREADONLY;
	filename->lpstrDefExt = defext;
	filename->nMaxFile = 32768;
	filename->lpstrFile = new WCHAR[filename->nMaxFile];
	if (filename->lpstrFile == NULL) {
		filename->nMaxFile = 0;
		if (alloc) delete filename, filename = NULL;
		return LIBSTATUS_SYSERR_NO_MEM;
	}
	LPWSTR buf = filename->lpstrFile;
	DWORD bufsz = filename->nMaxFile;
	SIZE_T pos = wcslen(defpathname);
	if (pos > 0) {
		wcscpy_s(buf, bufsz, defpathname);
		SIZE_T ipos = pos;
		do {
			if (filename->lpstrFile[ipos] == L'.') {
				if ((wcsicmp(&buf[ipos], L".m3d") == 0) ||
					(wcsicmp(&buf[ipos], L".a3d") == 0) ||
					(wcsicmp(&buf[ipos], L".t3d") == 0))
					pos = ipos;
				break;
			}
		} while(ipos--);
	} else {
		const WCHAR untitled[] = L"untitled";
		wcscpy_s(buf, bufsz, untitled);
		pos += _countof(untitled) - 1;
	}
	if (suffix) {
		wcscpy_s(&buf[pos], bufsz - pos, suffix);
		pos += wcslen(suffix);
	}
	wcscpy_s(&buf[pos], bufsz - pos, L".");
	wcscpy_s(&buf[++pos], bufsz - pos, defext);
	BOOL ret;
	if (save)
		ret = GetSaveFileName(filename);
	else
		ret = GetOpenFileName(filename);
	if (!ret) {
		if (alloc)
			Utils::FreeFileName(filename);
		return LIBSTATUS_ERR_ABORTED;
	}

	return LIBSTATUS_SUCCESS;
}

void Utils::FreeFileName(OPENFILENAME* & filename)
{
	if (filename == NULL)
		return;
	if (filename->lpstrFile != NULL)
		delete[] filename->lpstrFile, filename->lpstrFile = NULL;
	delete filename, filename = NULL;
}

UINT Utils::OpenCSVFile( std::wifstream & file, LPWSTR pathname )
{
	static const char bomU8[] = { '\xef', '\xbb', '\xbf' };
	static const char bomU16le[] = { '\xff', '\xfe' };
	static const char bomU16be[] = { '\xfe', '\xff' };

	std::ifstream bom;
	SetLastError(ERROR_SUCCESS);
	bom.open(pathname, std::ios::in|std::ios::binary);
	if (!bom)
		return LIBSTATUS_SYSERR | GetLastError();
	char bombuf[3];
	SetLastError(ERROR_SUCCESS);
	bom.read(bombuf, 3);
	if (!bom)
		return LIBSTATUS_SYSERR | GetLastError();
	bom.close();
	typedef enum {
		ASCII =   0x00, // 3 LSbs for BOM length, rest is unique identifier
		UTF8 =    0x0B,
		UTF16LE = 0x12,
		UTF16BE = 0x1A
	} ENC;
	ENC enc = ASCII;
	if (memcmp(bombuf, bomU8, 3) == 0) enc = UTF8;
	else if (memcmp(bombuf, bomU16le, 2) == 0) enc = UTF16LE;
	else if (memcmp(bombuf, bomU16be, 2) == 0) enc = UTF16BE;
	switch(enc) {
	case ASCII: file.imbue(std::locale(".1251")); break;
	case UTF8: file.imbue(std::locale(file.getloc(), new std::codecvt_utf8<wchar_t, 0x10ffff, std::consume_header>)); break;
	case UTF16LE: file.imbue(std::locale(file.getloc(), new std::codecvt_utf16<wchar_t, 0x10ffff, std::codecvt_mode(std::consume_header|std::little_endian)>)); break;
	case UTF16BE: file.imbue(std::locale(file.getloc(), new std::codecvt_utf16<wchar_t, 0x10ffff, std::consume_header>)); break;
	}
	SetLastError(ERROR_SUCCESS);
	file.open(pathname, std::ios::in|std::ios::binary);
	if (!file)
		return LIBSTATUS_SYSERR | GetLastError();

	return ERROR_SUCCESS;
}

/* from iproute2 */
INT Utils::Matches( LPCWSTR cmd, LPCWSTR pattern )
{
	size_t len = wcslen(cmd);
	if (len > wcslen(pattern))
		return -1;
	return wmemcmp(pattern, cmd, len);
}

// from Notepad2
/**
 * If the character is an hexa digit, get its value.
 */
static INT GetHexDigit( WCHAR ch ) {
	if (ch >= '0' && ch <= '9') {
		return ch - '0';
	}
	if (ch >= 'A' && ch <= 'F') {
		return ch - 'A' + 10;
	}
	if (ch >= 'a' && ch <= 'f') {
		return ch - 'a' + 10;
	}
	return -1;
}

/**
 * Convert C style \a, \b, \f, \n, \r, \t, \v, \xhh and \uhhhh into their indicated characters.
 */
UINT Utils::UnSlash( LPWSTR s ) {
	LPWSTR sStart = s;
	LPWSTR o = s;

	while (*s) {
		if (*s == '\\') {
			s++;
			if (*s == 'a')
				*o = '\a';
			else if (*s == 'b')
				*o = '\b';
			else if (*s == 'f')
				*o = '\f';
			else if (*s == 'n')
				*o = '\n';
			else if (*s == 'r')
				*o = '\r';
			else if (*s == 't')
				*o = '\t';
			else if (*s == 'v')
				*o = '\v';
			else if (*s == 'x' || *s == 'u') {
				BOOL bShort = (*s == 'x');
				WCHAR val;
				INT hex;
				val = 0;
				hex = GetHexDigit(*(s+1));
				if (hex >= 0) {
					s++;
					val = hex;
					hex = GetHexDigit(*(s+1));
					if (hex >= 0) {
						s++;
						val *= 16;
						val += hex;
						if (!bShort) {
							hex = GetHexDigit(*(s+1));
							if (hex >= 0) {
								s++;
								val *= 16;
								val += hex;
								hex = GetHexDigit(*(s+1));
								if (hex >= 0) {
									s++;
									val *= 16;
									val += hex;
								}
							}
						}
					}
					if (val)
						*o = val;
					else
						o--;
				}
				else
					o--;
			}
			else
				*o = *s;
		}
		else
			*o = *s;
		o++;
		if (*s) {
			s++;
		}
	}
	*o = '\0';
	return (UINT)(o - sStart);
}

BOOL Utils::ComStrStatus( CString & str, UINT status, BOOL append )
{
	if (status & LIBSTATUS_SYSERR) { // LIBSTATUS_SYSERR is a single bit flag indicating system error number coexistence, treat it specially
		LPTSTR lpBuffer = NULL;
		DWORD num;
		num = FormatMessage(
				FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL,
				status & ~LIBSTATUS_SYSERR, // drop the flag
				MAKELANGID(LANG_NEUTRAL,SUBLANG_DEFAULT),
				(LPTSTR)&lpBuffer,
				0,
				NULL);
		if (num > 0) {
			if (append)
				str.Append(lpBuffer);
			else
				str.SetString(lpBuffer);
		}
		if (lpBuffer)
			LocalFree(lpBuffer);
		return TRUE;
	}
	// it is not a system error, check for common exclusive messages
	if (status & LIBSTATUS_EXCLUSIVE) {
		LPCWSTR strptr = NULL;
		switch(status) {
		case LIBSTATUS_ERR_ABORTED:				strptr = L"Операция отменена";								break;
		case LIBSTATUS_ERR_API:					strptr = L"Фатальная ошибка API";							break;
		case LIBSTATUS_ERR_INVARG:				strptr = L"Недействительный аргумент";						break;
		case LIBSTATUS_ERR_UNKNOWN:				strptr = L"Неизвестная ошибка";								break;
		}
		if (strptr == NULL)
			return FALSE; // status contains no common errors
		if (append)
			str.Append(strptr);
		else
			str.SetString(strptr);
		return TRUE;
	}

	return FALSE;
}
