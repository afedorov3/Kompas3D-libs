#ifndef _UTILS_H
#define _UTILS_H

namespace Utils {

	const LPCWSTR CSVfilter = L"CSV פאיכû\0*.csv;*.txt\0ֲסו פאיכû\0*\0";
	const LPCWSTR TXTfilter = L"TXT פאיכû\0*.txt\0ֲסו פאיכû\0*\0";
	const LPCWSTR CSVext = L"csv";
	const LPCWSTR TXText = L"txt";

	void SanitizeString( LPWSTR str, LPCWSTR notallow, WCHAR replace );
	BOOL IsValidGUID( LPCWSTR guidStr, BOOL embraced = TRUE );
	INT ParsePath( LPCWSTR path, SIZE_T len, LPWSTR * endptr = NULL );
	BOOL IsPathRelative( INT ppath );
	BOOL IsPathRelative( LPCWSTR path );
	BOOL IsPathNetwork( INT ppath );
	BOOL IsPathNetwork( LPCWSTR path );
	BOOL IsPathDevice( INT ppath );
	BOOL IsPathDevice( LPCWSTR path );
	BOOL IsPathPipe( INT ppath );
	BOOL IsPathPipe( LPCWSTR path );
	BOOL SanitizeFileSystemString( LPWSTR str, SIZE_T strlen, WCHAR replacement, BOOL filename = FALSE, BOOL notrailsep = FALSE, LPWSTR * rootptr = NULL );
	BOOL SanitizeFileSystemString( CString & str, WCHAR replacement, BOOL filename = FALSE, BOOL notrailsep = FALSE, LPWSTR * rootptr = NULL );
	LPWSTR AbsPath( LPCWSTR path, LPCWSTR reference, BOOL reflock = FALSE, WCHAR sareplacement = '_' );
	BOOL CreateDirectoryRecursively( const CString directory );
	DWORD GetModulePathName( LPWSTR pathname, SIZE_T pnsz, SIZE_T *nameidx = NULL );
	BOOL Str2Clipboard(LPCWSTR str);

	BOOL ParseDouble( LPCWSTR str, LPCWSTR & endptr, DOUBLE & val, _locale_t locale = NULL );
	double RadtoDeg(double Rad);
	double DegtoRad(double Deg);

	UINT GetFileName( BOOL save, LPCWSTR defpathname, LPCWSTR title, LPCWSTR suffix, LPCWSTR filter, LPCWSTR defext, OPENFILENAME* & filename );
	void FreeFileName(OPENFILENAME* & filename);
	UINT OpenCSVFile( std::wifstream & file, LPWSTR pathname );

	INT Matches( LPCWSTR cmd, LPCWSTR pattern );
	UINT UnSlash( LPWSTR s ); 	// from Notepad2

	BOOL ComStrStatus( CString & str, UINT status, BOOL append );
} /* Utils */

#endif /* _UTILS_H */
