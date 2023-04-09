////////////////////////////////////////////////////////////////////////////////
//
// 3Dutils.cpp
//
////////////////////////////////////////////////////////////////////////////////
#include "stdafx.h"

#include "3Dutils.h"
#include "ColorChangeDlg.h"
#include "SelectParmDlg.h"
#include "OrientDlg.h"
#include "..\common\OverwriteDlg\OverwriteDlg.h"
#include "..\common\utils.h"
#include "..\common\CSVParser.h"
#include "..\common\CSVWriter.h"

#include "version.h"

#define _USE_MATH_DEFINES
#include <math.h>

#include <codecvt>

#define V_ISDISPATCH(X)     (V_VT(X)&VT_DISPATCH)

PALETTEENTRY DEFPALETTE[] = {
	{40, 40, 40, 0},	// black plastic
	{211, 212, 213, 0},	// tin
	{209, 179, 96, 0},	// gold light
	{181, 182, 181, 0},	// nickel
	{106, 102, 101, 0},	// black plastic marking
	{60, 60, 60, 0},	// metal marking black paint
	{30, 30, 30, 0},	// black paint
	{184, 115, 51, 0},	// copper
	{130, 70, 17},		// copper wire yellow enamelled
	{203, 205, 206, 0},	// aluminium
	{160, 149, 63, 0},	// brass
	{227, 222, 219, 0},	// chrome
	{202, 204, 206, 0},	// stainless steel (Crown Diamond 7407-21)
	{206, 172, 70, 0},	// gold
	{200, 157, 30, 0},	// gold dark
	{250, 250, 250, 0},	// white
};

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;

typedef enum {
	TYPE_PTR,
	TYPE_LONG,
	TYPE_ULONG,
	TYPE_DOUBLE,
	TYPE_STR,
	TYPE_BSTR
} VALTYPE;

#define showval(type, val) showVAL(L#val, type, val)
static void showVAL( LPCWSTR caption, VALTYPE type, ... )
{
	va_list args;
	CString str;

	va_start(args, type);
	switch (type) {
	case TYPE_PTR:
		str.Format(L"%p", va_arg(args, void*));
		break;
	case TYPE_LONG:
		str.Format(L"%d", va_arg(args, LONG));
		break;
	case TYPE_ULONG:
		str.Format(L"%u", va_arg(args, ULONG));
		break;
	case TYPE_DOUBLE:
		str.Format(L"%f", va_arg(args, DOUBLE));
		break;
	case TYPE_STR:
		str = va_arg(args, LPCWSTR);
		break;
	case TYPE_BSTR:
		str.Append(va_arg(args, _bstr_t));
		break;
	default:
		va_end(args);
		return;
	}
	va_end(args);

	MessageBox(NULL, str, caption, MB_OK);
}

#include <time.h>
typedef signed int clockid_t;

inline struct tm *localtime_r(const time_t *timep, struct tm *result)
{
	if (localtime_s(result, timep) == 0)
		return result;
	return NULL;
}

struct timespec {
	time_t   tv_sec;        /* seconds */
	long     tv_nsec;       /* nanoseconds */
};

/* only CLOCK_REALTIME currently implemented */
#define CLOCK_REALTIME 0
static int clock_gettime(clockid_t clk_id, struct timespec *tp) {
	union {
		FILETIME ft;
    	ULARGE_INTEGER ul;
	} t;
	/* FILETIME of Jan 1 1970 00:00:00. */
	static const unsigned __int64 epoch = 
		((unsigned __int64) 116444736000000000ULL);

	if (clk_id != CLOCK_REALTIME) {
		errno = EINVAL;
		return -1;
	}

	GetSystemTimeAsFileTime(&t.ft);
	t.ul.QuadPart -= epoch;
	tp->tv_sec = (time_t) (t.ul.QuadPart / 10000000L);
	tp->tv_nsec = (long) ((t.ul.QuadPart % 10000000L) * 100);

	return 0;	
}

/* from strace */
static void ts_sub(struct timespec *ts, const struct timespec *a,
								 const struct timespec *b)
{
	ts->tv_sec = a->tv_sec - b->tv_sec;
	ts->tv_nsec = a->tv_nsec - b->tv_nsec;
	if (ts->tv_nsec < 0) {
		ts->tv_sec--;
		ts->tv_nsec += 1000000000L;
	}
}

/* from strace */
static void addtimestamp(CString &str, int lvl)
{
	struct timespec ts;

	if (lvl <= 0)
		return;
	clock_gettime(CLOCK_REALTIME, &ts);
	if (lvl > 3) {	/* relative.mS */
		static struct timespec ots;
		struct timespec td;

		if (ots.tv_sec == 0)
			ots = ts;
		ts_sub(&td, &ts, &ots);
		str.AppendFormat(L"%6ld.%06ld",
			(long) td.tv_sec, td.tv_nsec / 1000L);
		ots = ts;
	} else if (lvl > 2) {
		str.AppendFormat(L"%lld.%06ld  ",
			(long long) ts.tv_sec, ts.tv_nsec / 1000L);
	} else {
		wchar_t strtime[sizeof("HH:MM:SS")];
		struct tm local;

		localtime_r(&ts.tv_sec, &local);
		wcsftime(strtime, _countof(strtime), L"%H:%M:%S", &local);
		if (lvl > 1)
			str.AppendFormat(L"%ls.%06ld  ",
				strtime, ts.tv_nsec / 1000L);
		else
			str.AppendFormat(L"%ls  ", strtime);
	}
}
#endif /* _DEBUG */

C3DutilsApp theApp;

C3DutilsApp::C3DutilsApp()
{
	m_newKompasAPI = NULL;
	m_kompas = NULL;
	m_numlocale = _wcreate_locale(LC_NUMERIC, L"C");

	m_selfilter = OM_BELONGSANY | OM_USECOLORMASK;
	m_selcolorovr = FALSE;
	m_selcolor = DEFCOLOR;
	m_seltolerance = CT_EXACT;

	m_oripreview = FALSE;
	m_oriautoplace = OrientDlg::AUTOPLACE_DEFAULT;
	m_oriposX = LONG_MIN;
	m_oriposY = LONG_MIN;

	m_chusecolor = useColorOur;
	m_chcolor = DEFCOLOR;
}

C3DutilsApp::~C3DutilsApp()
{
	_free_locale(m_numlocale);
}

BOOL C3DutilsApp::InitInstance()
{
	// InitCommonControls() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	InitCommonControls();

	AfxEnableControlContainer();

	LoadOrients();

	return CWinApp::InitInstance();
}

BOOL C3DutilsApp::ExitInstance()
{
	m_newKompasAPI = NULL;
	m_kompas = NULL;

	return CWinApp::ExitInstance();
}

// public
// document management
ksAPI7::IKompasDocument3DPtr C3DutilsApp::GetCurrentDocument()
{
	ksAPI7::IKompasDocument3DPtr doc3d = NULL;

	// get Kompas API
	if (!hasAPI())
		GetKompas();

	// Get active document
	if (hasAPI())
		doc3d = m_newKompasAPI->GetActiveDocument();

	return doc3d;
}

// faces
BOOL C3DutilsApp::FaceSelectSameColor( ksAPI7::IKompasDocument3DPtr & doc3d, _variant_t & selection, ULONG &selcount, COLORREF color, ULONG tolerance, ULONG filter, ksAPI7::IModelObjectPtr objptr )
{
	ksAPI7::IPart7Ptr srcpart = NULL;
	if (objptr != NULL)
		srcpart = objptr->Part;

	ksAPI7::IPart7Ptr toppart( doc3d->GetTopPart());
	if (toppart == NULL)
		return FALSE;

	if (srcpart != NULL) {
		switch(filter & OM_BELONGSMASK) {
		case OM_BELONGSPART:
		case OM_BELONGSBODY:
			return FaceSelectSameColor(doc3d, selection, selcount, srcpart, color, tolerance, filter, objptr);
		}
	}

	ksAPI7::IParts7Ptr parts(toppart->Parts);
	BOOL ret = TRUE;

	if (parts != NULL) {
		for(LONG pindex = 0; pindex < parts->Count; pindex++)
			if (!FaceSelectSameColor(doc3d, selection, selcount, parts->Part[pindex], color, tolerance, filter, objptr))
				ret = FALSE;
	}
	if (!FaceSelectSameColor(doc3d, selection, selcount, toppart, color, tolerance, filter, objptr)) {
		ret = FALSE;
	}

	return ret;
}

BOOL C3DutilsApp::FaceSelectSameColor( ksAPI7::IKompasDocument3DPtr & doc3d, _variant_t & selection, ULONG &selcount, ksAPI7::IPart7Ptr & part, COLORREF color, ULONG tolerance, ULONG filter, ksAPI7::IModelObjectPtr objptr )
{
	ksAPI7::IPart7Ptr srcpart = NULL;
	if (objptr != NULL)
		srcpart = objptr->Part;
	ksAPI7::IFeature7Ptr srcowner = NULL;
	if (objptr != NULL)
		srcowner = objptr->Owner;

	ksAPI7::IModelContainerPtr mc(part);
	if (mc == NULL)
		return FALSE;

	_variant_t objects(mc->Objects[o3d_face]);
	if (!V_ISDISPATCH(&objects))
		return FALSE;

	LPDISPATCH HUGEP * pObjs = NULL;
	HRESULT hr = E_FAIL;
	ULONG objcount = 0;
	BOOL ret = FALSE;

	if (V_ISARRAY(&objects)) {
		objcount = V_ARRAY(&objects)->rgsabound[0].cElements - V_ARRAY(&objects)->rgsabound[0].lLbound;
		hr = ::SafeArrayAccessData( V_ARRAY(&objects), (void HUGEP* FAR*)&pObjs );
	} else {
		objcount = 1;
		pObjs = &V_DISPATCH(&objects);
		hr = S_OK;
	}
	if (!FAILED(hr) && pObjs) {
		for (ULONG index = 0; index < objcount; index++) {
			ksAPI7::IModelObjectPtr mobj(pObjs[index]);
			if ((mobj == NULL) || (mobj->ModelObjectType != o3d_face))
				continue;

			if (mobj == objptr) // skip source
				continue;

			switch(filter & OM_BELONGSMASK) {
			case OM_BELONGSPART: // same part only
				if (mobj->Part != srcpart)
					continue;
				break;
			case OM_BELONGSBODY: // same body only
				if (mobj->Owner != srcowner)
					continue;
				break;
			}

			ksAPI7::IColorParam7Ptr colorparm(mobj);
			if (colorparm == NULL)
				continue;

			if (filter & OM_USECOLORMASK) {
				switch(colorparm->UseColor) {
				case useColorOur:
					if ((filter & OM_USECOLORCUSTOM) == 0)
						continue;
					break;
				case useColorOwner:
					if ((filter & OM_USECOLOROWNER) == 0)
						continue;
					break;
				case useColorSource:
					if ((filter & OM_USECOLORSOURCE) == 0)
						continue;
					break;
				}
			}

			if (ColorDifference(colorparm->Color, color) > tolerance)
				continue;

			if (SelectObjAdd(selection, pObjs[index]))
				selcount++;
		}

		ret = TRUE;
	}
	if (V_ISARRAY(&objects) && pObjs)
		::SafeArrayUnaccessData( V_ARRAY(&objects) );

	return ret;
}

BOOL C3DutilsApp::SelectDialog( ksAPI7::IKompasDocument3DPtr & doc3d )
{
	ksAPI7::IModelObjectPtr source = NULL;
	COLORREF srccol = DEFCOLOR;

	ksAPI7::ISelectionManagerPtr selmgr(doc3d->SelectionManager);
	if (selmgr == NULL)
		return FALSE;
	_variant_t selection = selmgr->SelectedObjects;

	if (V_ISDISPATCH(&selection)) {
		LPDISPATCH HUGEP * pSelObjs = NULL;
		HRESULT hr = E_FAIL;
		ULONG count = 0;
		if (V_ISARRAY(&selection)) {
			count = V_ARRAY(&selection)->rgsabound[0].cElements - V_ARRAY(&selection)->rgsabound[0].lLbound;
			hr = ::SafeArrayAccessData( V_ARRAY(&selection), (void HUGEP* FAR*)&pSelObjs );
		} else {
			count = 1;
			pSelObjs = &V_DISPATCH(&selection);
			hr = S_OK;
		}
		if (!FAILED(hr) && pSelObjs) {
			for (ULONG index = 0; index < count; index++) {
				ksAPI7::IKompasAPIObjectPtr apiobj(pSelObjs[index]);
				ULONG type = apiobj->Type;
				if ((apiobj == NULL)/* || apiobj->Type != ksObjectModelObject*/) // is model object
					continue;

				ksAPI7::IModelObjectPtr mobj(apiobj);
				if (mobj == NULL)
					continue;

				ksAPI7::IColorParam7Ptr colorparm(mobj);	// has color param
				if (colorparm == NULL)
					continue;

				source = mobj;
				srccol = colorparm->Color;
				break;
			}
		}
		if (V_ISARRAY(&selection) && pSelObjs)
			::SafeArrayUnaccessData( V_ARRAY(&selection) );
	}

	SelectParmDlg Dlg(m_selfilter,
		source != NULL,	// has source
		FALSE, // colorovr
		(source == NULL)?m_selcolor:srccol, // color
		m_seltolerance); // tolerance
	if (Dlg.DoModal() == IDOK) {
		m_selfilter = Dlg.GetFilter();
		m_selcolorovr = Dlg.GetColorOverride();
		if (m_selcolorovr || (source == NULL))
			m_selcolor = Dlg.GetColor();
		m_seltolerance = Dlg.GetTolerance();

		ULONG selcount = 0;
		if (FaceSelectSameColor(doc3d, // document
				selection, // current selection
				selcount, // selected count
				(m_selcolorovr || (source == NULL))?m_selcolor:srccol, // color
				m_seltolerance, // tolerance
				m_selfilter, // filter
				source) // source
				&& selcount > 0) {
			SelectCommit(doc3d, selection);
			SelectClear(selection);
		}
	}

	return TRUE;
}

BOOL C3DutilsApp::FaceColorChange( ksAPI7::IKompasDocument3DPtr & doc3d, ULONG &chcount, ULONG usecolor, COLORREF color )
{
	BOOL ret = FALSE;

	ksAPI7::ISelectionManagerPtr selmgr(doc3d->SelectionManager);
	if (selmgr == NULL)
		return FALSE;
	_variant_t selection = selmgr->SelectedObjects;

	if (V_ISDISPATCH(&selection)) {
		LPDISPATCH HUGEP * pSelObjs = NULL;
		HRESULT hr = E_FAIL;
		ULONG count = 0;
		if (V_ISARRAY(&selection)) {
			count = V_ARRAY(&selection)->rgsabound[0].cElements - V_ARRAY(&selection)->rgsabound[0].lLbound;
			hr = ::SafeArrayAccessData( V_ARRAY(&selection), (void HUGEP* FAR*)&pSelObjs );
		} else {
			count = 1;
			pSelObjs = &V_DISPATCH(&selection);
			hr = S_OK;
		}
		if (!FAILED(hr) && pSelObjs) {
			for (ULONG index = 0; index < count; index++) {
				ksAPI7::IKompasAPIObjectPtr apiobj(pSelObjs[index]);
				if ((apiobj == NULL)/* || apiobj->Type != ksObjectModelObject*/) // is model object
					continue;

				ksAPI7::IModelObjectPtr mobj(apiobj);
				if ((mobj == NULL) || (mobj->ModelObjectType != o3d_face))
					continue;

				if (ObjSetColor( mobj, usecolor, color )) {
					if (mobj->Update())
						chcount++;
				}
			}
			ret = TRUE;
		}
		if (V_ISARRAY(&selection) && pSelObjs)
			::SafeArrayUnaccessData( V_ARRAY(&selection) );
	}

	return ret;
}

BOOL C3DutilsApp::ColorChangeDialog( ksAPI7::IKompasDocument3DPtr & doc3d )
{
	ColorChangeDlg Dlg(m_chusecolor, m_chcolor);
	if (Dlg.DoModal() == IDOK) {
		m_chusecolor = Dlg.GetUseColor();
		m_chcolor = Dlg.GetColor();

		ULONG chcount = 0;
		FaceColorChange(doc3d, chcount, m_chusecolor, m_chcolor);
	}

	return TRUE;
}

// selection array
BOOL C3DutilsApp::SelectObjAdd( _variant_t &selection, ksAPI7::IKompasAPIObjectPtr obj )
{
	LPDISPATCH HUGEP * psObjs = NULL;
	
	ULONG count = 0;
	BOOL ret = FALSE;

	if (V_ISDISPATCH(&selection)) {
		if (V_ISARRAY(&selection)) {
			count = V_ARRAY(&selection)->rgsabound[0].cElements - V_ARRAY(&selection)->rgsabound[0].lLbound;

			SAFEARRAYBOUND sabNewArray;
			sabNewArray.cElements = count + 1;
			sabNewArray.lLbound = 0;
			HRESULT hr = ::SafeArrayRedim( V_ARRAY(&selection), &sabNewArray);
			if (!FAILED(hr)) {
				LONG i = count;

				::SafeArrayPutElement( V_ARRAY(&selection), &i, obj);
				ret = TRUE;
			}
		} else {
			SAFEARRAYBOUND sabNewArray;
			sabNewArray.cElements = 2;
			sabNewArray.lLbound = 0;

			SAFEARRAY * pSafe = ::SafeArrayCreate( VT_DISPATCH, 1, &sabNewArray );
			if( pSafe ) {
				LONG i = 0;
				::SafeArrayPutElement( pSafe, &i, V_DISPATCH(&selection));
				i++;
				::SafeArrayPutElement( pSafe, &i, obj);
				V_DISPATCH(&selection) = NULL;
				V_VT(&selection) = VT_DISPATCH | VT_ARRAY;
				V_ARRAY(&selection) = pSafe;
				ret = TRUE;
			}
		}
	} else if (V_VT(&selection) == VT_EMPTY) {
		V_DISPATCH(&selection) = obj;
		V_VT(&selection) = VT_DISPATCH;

		ret = TRUE;
	}

	return ret;
}

BOOL C3DutilsApp::SelectClear( _variant_t & selection )
{
	if (V_ISARRAY(&selection) && (V_ARRAY(&selection) != NULL))
		SafeArrayDestroy(V_ARRAY(&selection));
	V_ARRAY(&selection) = NULL;
	V_DISPATCH(&selection) = NULL;
	V_VT(&selection) = VT_EMPTY;

	return TRUE;
}

BOOL C3DutilsApp::SelectCommit( ksAPI7::IKompasDocument3DPtr & doc3d, _variant_t &selection )
{
	ksAPI7::ISelectionManagerPtr selmgr(doc3d->SelectionManager);
	if (selmgr == NULL)
		return FALSE;

	return selmgr->Select(selection);
}

BOOL C3DutilsApp::PlacementDialog( ksAPI7::IKompasDocument3DPtr & doc3d )
{
	OrientDlg Dlg( doc3d, m_oripreview, m_oriautoplace, m_oriposX, m_oriposY );
	INT_PTR ret = Dlg.DoModal();
	m_oripreview = Dlg.GetPreview();
	m_oriautoplace = Dlg.GetAutoplace();
	Dlg.GetWinPos(m_oriposX, m_oriposY);

	return ret;
}

BOOL C3DutilsApp::VariablesExport( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment )
{
	BOOL vready = FALSE, clipbrd = FALSE;
	if (GetKeyState(VK_CONTROL) & 0x80) {
		vready = clipbrd = TRUE;
	}
	if (GetKeyState(VK_SHIFT) & 0x80) { clipbrd = !clipbrd; }

	ksAPI7::IEmbodimentPtr e = Document::EmbodimentPtrGet(doc3d, embodiment);
	if (e == NULL)
		return FALSE;
	ksAPI7::IPart7Ptr epart = e->GetPart();
	if (epart == NULL)
		return FALSE;
	ksAPI7::IFeature7Ptr feature(epart);
	if (feature == NULL)
		return FALSE;

	if (feature->GetVariablesCount(FALSE, FALSE) < 1) {
		MessageBox((HWND)GetHWindow(), L"Главный раздел исполнения не содержит переменных", L"Экспорт переменных", MB_OK|MB_ICONWARNING);
		return FALSE;
	}

	OPENFILENAME *filename = NULL;
	UINT lret;
	if (!clipbrd) {
		CString suf;
		Document::EmbodimentMarkingGet(doc3d, embodiment, ksVMEmbodimentNumber, &suf);
		suf += L"_variables";
		lret = Utils::GetFileName(TRUE, (LPCWSTR)doc3d->GetPathName(), vready?L"Экспорт переменных (строка)":L"Экспорт переменных", (LPCWSTR)suf,
			vready?Utils::TXTfilter:Utils::CSVfilter, vready?Utils::TXText:Utils::CSVext, filename);
		if (lret != LIBSTATUS_SUCCESS) {
			if (lret == LIBSTATUS_ERR_ABORTED)
				return TRUE;
			return FALSE;
		}
	}

	BOOL ret = FALSE;
	_variant_t vars(feature->GetVariables(FALSE, FALSE));
	if (V_ISDISPATCH(&vars)) {
		LPDISPATCH HUGEP * pVars = NULL;
		HRESULT hr = E_FAIL;
		ULONG count = 0;
		if (V_ISARRAY(&vars)) {
			count = V_ARRAY(&vars)->rgsabound[0].cElements - V_ARRAY(&vars)->rgsabound[0].lLbound;
			hr = ::SafeArrayAccessData( V_ARRAY(&vars), (void HUGEP* FAR*)&pVars );
		} else {
			count = 1;
			pVars = &V_DISPATCH(&vars);
			hr = S_OK;
		}
		if (!FAILED(hr) && pVars) {
			ULONG done = 0;
			CSVWriter csv(3); // name, value/expression, comment
			CString vstr;
			if (!vready) csv << L"Имя" << L"Выражение" << L"Комментарий";
			for (ULONG index = 0; index < count; index++) {
				ksAPI7::IVariable7Ptr var(pVars[index]);
				if (var == NULL)
					continue;

				if (vready) {
					LPWSTR eptr;
					wcstod((LPCWSTR)var->Expression, &eptr);
					if (*eptr != L'\0')
						vstr.AppendFormat(L"%s%s=\"%s\"", done?L";":L"", (LPCWSTR)var->Name, (LPCWSTR)var->Expression);
					else
						vstr.AppendFormat(L"%s%s=%g", done?L";":L"", (LPCWSTR)var->Name, var->Value);
				} else {
					csv << (LPCWSTR)var->Name;
					csv << (LPCWSTR)var->Expression;
					csv << (LPCWSTR)var->Note;
				}

				done++;
			}

			if (V_ISARRAY(&vars) && pVars)
				::SafeArrayUnaccessData( V_ARRAY(&vars) );

			if (done > 0) {
				CString str;
				SetLastError(ERROR_SUCCESS);
				if (vready) {
					if (clipbrd) {
						ret = Utils::Str2Clipboard((LPCWSTR)vstr);
					} else {
						std::wofstream file;

						file.imbue(std::locale(file.getloc(), new std::codecvt_utf8<wchar_t, 0x10ffff, std::consume_header>));
						file.open(filename->lpstrFile, std::ios::out);
						if (file) {
							file << (LPCWSTR)vstr << std::endl;
							DWORD err = GetLastError();
							if (file)
								ret = TRUE;
							file.close();
							SetLastError(err);
						}
					}
				} else {
					csv.newRow();
					if (clipbrd) {
						ret = Utils::Str2Clipboard(csv.toString().c_str());
					} else
						ret = csv.writeToFile(filename->lpstrFile);
				}

				if (ret) {
					str.Format(L"Экспортировано %d переменных в %s", done, clipbrd?L"буфер обмена":filename->lpstrFile);
					if (vready)
						str.AppendFormat(L"\nДлина строки %d символов", vstr.GetLength());
					MessageBox((HWND)GetHWindow(), (LPCWSTR)str, vready?L"Экспорт переменных [шаблон]":L"Экспорт переменных [список]", MB_OK|MB_ICONINFORMATION);
				} else {
					DWORD err = GetLastError();
					if (err != ERROR_SUCCESS)
						lret = LIBSTATUS_SYSERR | err;
					else
						lret = LIBSTATUS_ERR_UNKNOWN;
					str.Format(L"Ошибка экспорта в %s: ", clipbrd?L"буфер обмена":filename->lpstrFile);
					Utils::ComStrStatus(str, lret, TRUE);
					MessageBox((HWND)GetHWindow(), (LPCWSTR)str, vready?L"Экспорт переменных [шаблон]":L"Экспорт переменных [список]", MB_OK|MB_ICONERROR);
				}
			}
		}
	}

	Utils::FreeFileName(filename);

	return ret;
}

BOOL C3DutilsApp::VariablesImport( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment )
{
	BOOL ret = FALSE;
	std::wifstream csvstream;
	CString report;

	OPENFILENAME *filename = NULL;
	UINT lret = Utils::GetFileName( FALSE, (LPCWSTR)doc3d->GetPathName(), L"Импорт переменных", L"_variables", Utils::CSVfilter, Utils::CSVext, filename);
	if (lret != LIBSTATUS_SUCCESS) {
		if (lret == LIBSTATUS_ERR_ABORTED)
			return TRUE;
		return FALSE;
	}

	if ((lret = Utils::OpenCSVFile(csvstream, filename->lpstrFile)) == ERROR_SUCCESS) {
		aria::csv::CsvParser *parser = new aria::csv::CsvParser(csvstream);
		parser->delimiter(L';');
		LONG total = 0, done = 0, dstat = 0;
		UINT ovr = BST_INDETERMINATE;
		CString discarded;
		LPCWSTR reason = L"";
		CString status;

		ksAPI7::IProgressBarIndicatorPtr progress(m_newKompasAPI->GetProgressBarIndicator());

		for (auto& row : *parser) {
			SIZE_T fields = row.size();
			LPWSTR endptr = NULL;
			ksAPI7::IVariable7Ptr var;

			if ((fields < 2) || !Document::IsVariableNameValid(row[0].c_str())) // name, value/expression, comment
				continue;

			total++;

			if (progress != NULL) {
				status.Format(L"Импорт переменной %s (%d/%d)", row[0].c_str(), done, total);
				progress->SetText((LPCWSTR)status);
			}

			if (var = Document::EmbodimentVariable(doc3d, row[0].c_str(), embodiment)) {
				// variable exists
				if (ovr == BST_INDETERMINATE) {
					CString vardet;
					if (fields > 2)
						vardet.Format(L"%s\t%s\t%s", row[0].c_str(), row[1].c_str(), row[2].c_str());
					else
						vardet.Format(L"%s\t%s", row[0].c_str(), row[1].c_str());
					OverwriteDlg Dlg(L"3DUtils", OverwriteDlg::ENTITY_VAR, (LPCWSTR)vardet);
					INT_PTR nRet = Dlg.DoModal();
					switch(nRet) {
					case IDNONE:
						ovr = BST_UNCHECKED;
					case IDNO:
						reason = L"переменная существует";
						goto discard;
					case IDALL:
						ovr = BST_CHECKED;
					case IDYES:
						break;
					default:
						total--;
						report.Format(L"Операция прервана, импортировано %d/%d переменных", done, total);
						goto abort;
					}
				} else if (ovr != BST_CHECKED) {
					reason = L"переменная существует";
					goto discard;
				}
			} else
				var = Document::EmbodimentVariable(doc3d, row[0].c_str(), embodiment, TRUE);

			if (var == NULL) {
				reason = L"ошибка создания переменной";
				goto discard;
			}

			if (fields > 2)
				var->Note = row[2].c_str();
			else
				var->Note = L"";
			SIZE_T explen = wcslen(row[1].c_str());
			if (explen < 1)
				var->Value = 1.0; // default value
			else if (!Document::VariableExprSet(var, row[1].c_str())) {
				// Try to parse as a link
				LPWSTR expr = new WCHAR[explen + 1];
				if (expr == NULL) {
					reason = L"нехватка памяти";
					goto discard;
				}
				wcscpy_s(expr, explen + 1, row[1].c_str());
				LPWSTR eptr = expr;
				LPCWSTR doc = NULL;
				LPCWSTR emb = NULL;
				LPCWSTR lvar = NULL;
				WCHAR ch;
				BOOL lset = FALSE;
				while (ch = *eptr) {
					if (ch == L'|') {
						*eptr = '\0';
						if (emb == NULL)
							emb = eptr + 1;
						else if (lvar == NULL)
							lvar = eptr + 1;
						else {
							delete[] expr;
							reason = L"неверный формат ссылки";
							goto discard;
						}
					}
					eptr++;
				}
				if (emb == NULL) {
					delete[] expr;
					reason = L"неверный формат ссылки";
					goto discard;
				}
				if (lvar == NULL) { // different embodimemt of the same document
#if 0
					// FIXME This wouldn't work until SetLink will be capable of accepting embodiment specification
					lvar = emb;
					emb = expr;
					doc = (LPCWSTR)doc3d->GetPathName() // L"", NULL or what? New unsaved document path is undefined
#else
					delete[] expr;
					reason = L"локальная ссылка не поддерживается";
					goto discard;
#endif
				} else
					doc = expr;
				if (!Document::IsVariableNameValid(lvar)) {            // check variable name
					delete[] expr;
					reason = L"неверное имя целевой переменной ссылки";
					goto discard;
				}
				if (doc && Utils::ParsePath(doc, emb - doc - 1) < 0) { // and document path if found FIXME Absolute path check?
					delete[] expr;
					reason = L"неверный путь к целевому файлу ссылки";
					goto discard;
				}
				lset = var->SetLink(doc, lvar);
				delete[] expr;
				if (!lset) {
					reason = L"ошибка создания ссылки";
					goto discard;
				}
			}

			done++;
			continue;
discard:
			if (dstat < 10) {
				if (fields > 2)
					discarded.AppendFormat(L"  %s\t%s\t%s\t%s\n", row[0].c_str(), row[1].c_str(), row[2].c_str(), reason);
				else
					discarded.AppendFormat(L"  %s\t%s\t%s\n", row[0].c_str(), row[1].c_str(), reason);
				reason = L"";
				++dstat;
			}
			continue;
		}

		if (progress != NULL)
			progress->SetText(L"Готово");

		if (total > 0) {
			ret = (total == done);

			report.Format(L"Импортировано %d/%d переменных", done, total);
abort:
			if (!discarded.IsEmpty()) {
				report += L"\nБыли отвергнуты следующие переменные:\n";
				report += discarded;
				dstat = total - done - dstat;
				if (dstat > 0)
					report.AppendFormat(L"... и еще %d", dstat);
			}
			report += L"\n\nПерестройте документ (F5) для обновления списка";
			MessageBox((HWND)GetHWindow(), (LPCWSTR)report, L"Импорт переменных", MB_OK|(ret ? MB_ICONINFORMATION : MB_ICONWARNING));
		} else
			MessageBox((HWND)GetHWindow(), L"В файле не найдено ни одного описания переменных", L"Импорт переменных", MB_OK|MB_ICONWARNING);
	} else {
		report.Format(L"Ошибка чтения файла %: ", filename->lpstrFile);
		Utils::ComStrStatus(report, lret, TRUE);
		MessageBox((HWND)GetHWindow(), (LPCWSTR)report, L"Импорт переменных", MB_OK|MB_ICONERROR);
	}

	Utils::FreeFileName(filename);

	return ret;
}

// utility
//----------------------------------------------------------------------------
/*
   Colour Difference Formula

   The following is the formula suggested by the W3C to determine
   the difference between two colours.

     (maximum (Red   value 1, Red   value 2) - minimum (Red   value 1, Red   value 2))
   + (maximum (Green value 1, Green value 2) - minimum (Green value 1, Green value 2))
   + (maximum (Blue  value 1, Blue  value 2) - minimum (Blue  value 1, Blue  value 2))

   The difference between the background colour and the foreground
   colour should be greater than 500.
 */
//----------------------------------------------------------------------------

UINT C3DutilsApp::ColorDifference( COLORREF c1, COLORREF c2 )
{
   return (max(GetRValue(c1), GetRValue(c2)) - min(GetRValue(c1), GetRValue(c2)))
        + (max(GetGValue(c1), GetGValue(c2)) - min(GetGValue(c1), GetGValue(c2)))
        + (max(GetBValue(c1), GetBValue(c2)) - min(GetBValue(c1), GetBValue(c2)));
}

BOOL C3DutilsApp::ObjSetColor( ksAPI7::IModelObjectPtr objptr, ULONG usecolor, COLORREF color )
{
	if (objptr == NULL)
		return FALSE;

	ksAPI7::IColorParam7Ptr colorparm(objptr);	// has color param
	if (colorparm == NULL)
		return FALSE;

	BOOL ret = FALSE;
	if (colorparm->UseColor != usecolor) {
		for(INT i = 0; i < 3; i++) { // some bug in setting usecolor? seems to set only on second assignment
			colorparm->UseColor = (ksUseColorEnum)usecolor;
			if (colorparm->UseColor == usecolor) {
				ret = TRUE;
				break;
			}
		}
	}

	if ((usecolor == useColorOur) && (colorparm->UseColor == useColorOur) && (colorparm->Color != color)) {
		for(INT i = 0; i < 3; i++) { // just in case
			colorparm->Color = color;
			if (colorparm->Color == color) {
				ret = TRUE;
				break;
			}
		}
	}

	return ret;
}

CPalette *C3DutilsApp::CreateDefaultPalete() {
	CPalette *defpalette = new CPalette;
	if (defpalette == NULL)
		return NULL;

	LOGPALETTE lp = { 0x300, 1, { 0, 0, 0, 0 } };
	if (!defpalette->CreatePalette(&lp))
		goto clean;

	SIZE_T entries = sizeof(DEFPALETTE) / sizeof(DEFPALETTE[0]);
	if (!defpalette->ResizePalette(entries))
		goto clean;

	if (defpalette->SetPaletteEntries(0, entries, DEFPALETTE) != entries)
		goto clean;

	return defpalette;
clean:
	delete defpalette;
	return NULL;
}

BOOL C3DutilsApp::SetPlacementEulerDeg( IPlacement *place, double rotXangle, double rotYangle, double rotZangle )
{
	_variant_t matrix;
	place->GetMatrix3D(&matrix);
	if (V_VT(&matrix) != ( VT_ARRAY | VT_R8 ))
		return FALSE;

	double HUGEP *pM = NULL;
	::SafeArrayAccessData( V_ARRAY(&matrix), (void HUGEP* FAR*)&pM );
	if (pM == NULL)
		return FALSE;

	SetEulerDegXYZ(pM, rotXangle, rotYangle, rotZangle);

	::SafeArrayUnaccessData( V_ARRAY(&matrix) );

	place->InitByMatrix3D(matrix);

	return TRUE;
}

BOOL C3DutilsApp::EulerDegfromPlacement( IPlacement *place, double& rotXangle, double& rotYangle, double& rotZangle )
{
	_variant_t matrix;
	place->GetMatrix3D(&matrix);
	if (V_VT(&matrix) != ( VT_ARRAY | VT_R8 ))
		return FALSE;

	double HUGEP *pM = NULL;
	::SafeArrayAccessData( V_ARRAY(&matrix), (void HUGEP* FAR*)&pM );
	if (pM == NULL)
		return FALSE;

	ExtractEulerRadXYZ(pM, rotXangle, rotYangle, rotZangle);

	::SafeArrayUnaccessData( V_ARRAY(&matrix) );

	rotXangle = Utils::RadtoDeg(rotXangle);
	rotYangle = Utils::RadtoDeg(rotYangle);
	rotZangle = Utils::RadtoDeg(rotZangle);

	return TRUE;
}

void C3DutilsApp::ExtractEulerRadXYZ(double* rotmatrix4x4, double& rotXangle, double& rotYangle, double& rotZangle)
{
	rotXangle = atan2(-rotmatrix4x4[6], rotmatrix4x4[10]);
	double cosY = sqrt(pow(rotmatrix4x4[0], 2) + pow(rotmatrix4x4[1], 2));
	rotYangle = atan2(rotmatrix4x4[2], cosY);
	double sinX = sin(rotXangle);
    double cosX = cos(rotXangle);
	rotZangle = atan2(cosX * rotmatrix4x4[4] + sinX * rotmatrix4x4[8], cosX * rotmatrix4x4[5] + sinX * rotmatrix4x4[9]);
}

void C3DutilsApp::SetEulerDegXYZ(double* rotmatrix4x4, double rotXangle, double rotYangle, double rotZangle)
{
	double sinX = SinD(rotXangle);
	double cosX = CosD(rotXangle);
	double sinY = SinD(rotYangle);
	double cosY = CosD(rotYangle);
	double sinZ = SinD(rotZangle);
	double cosZ = CosD(rotZangle);

	rotmatrix4x4[0] =  cosY*cosZ;                   rotmatrix4x4[1] = -cosY*sinZ;                   rotmatrix4x4[2] =  sinY;
	rotmatrix4x4[4] =  sinX*sinY*cosZ + cosX*sinZ;  rotmatrix4x4[5] = -sinX*sinY*sinZ + cosX*cosZ;  rotmatrix4x4[6] = -sinX*cosY;
	rotmatrix4x4[8] = -cosX*sinY*cosZ + sinX*sinZ;  rotmatrix4x4[9] =  cosX*sinY*sinZ + sinX*cosZ;  rotmatrix4x4[10] = cosX*cosY;
}

CString C3DutilsApp::LoadStr( UINT strID )
{
	CString str;
	str.LoadStringW(m_hInstance, strID);
	return str;
}

SIZE_T C3DutilsApp::LoadStr( UINT strID, LPWSTR buf, SIZE_T bufsz )
{
	return LoadString(m_hInstance, strID, buf, bufsz);
}

void C3DutilsApp::EnableTaskAccess( BOOL enable )
{
	m_kompas->ksEnableTaskAccess(enable);
}

// private
BOOL C3DutilsApp::GetKompas() 
{
	if (hasAPI()) return TRUE;

	const SIZE_T mpsz = 32768;
	LPWSTR modulepath = new WCHAR[mpsz];
	if (modulepath == NULL)
		goto noapi;

	DWORD flen = ::GetModuleFileNameW(NULL, modulepath, mpsz);
	if ((flen == 0) || (flen >= mpsz))
		goto noapi;

	size_t basename;
	for(basename = flen; (basename > 0) && (modulepath[basename - 1] != '\\'); basename--);
	if (!LoadStr(IDR_API5, modulepath + basename, mpsz - basename))
		goto noapi;

	HINSTANCE hAppAuto = LoadLibraryW(modulepath); // идентификатор kAPI5.dll
	if (hAppAuto == NULL)
		goto noapi;
	delete[] modulepath, modulepath = NULL;

	typedef LPDISPATCH ( WINAPI * FCreateKompasObject )(); 
	FCreateKompasObject pCreateKompasObject = 
		(FCreateKompasObject)GetProcAddress(hAppAuto, "CreateKompasObject");	
	if (pCreateKompasObject) 
		m_kompas = IDispatchPtr( pCreateKompasObject(), false/*AddRef*/ );
	FreeLibrary(hAppAuto);

	if (m_kompas == NULL)
		goto noapi;
	m_newKompasAPI = m_kompas->ksGetApplication7();

	return TRUE;

noapi:
	DWORD dwErr = GetLastError();
	if (modulepath != NULL) delete[] modulepath;
	CString err = theApp.LoadStr(IDS_NOAPI);
	err += L": ";
	Utils::ComStrStatus(err, LIBSTATUS_SYSERR | dwErr, TRUE);

	MessageBox((HWND)GetHWindow(), err, theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONERROR);

	return FALSE;
}

UINT C3DutilsApp::LoadOrients()
{
	enum {
		FIELD_ANGLEX = 0,
		FIELD_ANGLEY,
		FIELD_ANGLEZ,
		FIELD_NAME,
		FIELD_TOOLTIP
	};
	const SIZE_T bufsz = 32768;
	LPWSTR mpath = NULL;
	aria::csv::CsvParser *parser = NULL;
	std::wifstream csvstream;

	mpath = new WCHAR[bufsz];
	if (mpath == NULL)
		return LIBSTATUS_SYSERR_NO_MEM;

	SIZE_T nameidx = 0;
	if (!Utils::GetModulePathName(mpath, bufsz - _countof(ORIENT_FILE) + 1, &nameidx)) {
		delete[] mpath;
		return FALSE;
	}

	wcscpy_s(&mpath[nameidx], bufsz - nameidx, ORIENT_FILE);

	if (Utils::OpenCSVFile(csvstream, mpath) != ERROR_SUCCESS)
		return FALSE;

	orient_t neworient;
	parser = new aria::csv::CsvParser(csvstream);
	parser->delimiter(L';');
	m_orients.clear();
	int count = 0;
	for (auto& row : *parser) {
		if (count >= MAX_ORIENTS)
			break;

		SIZE_T fields = row.size();
		if (fields < 4)
			continue;

		if (wcslen(row[FIELD_NAME].c_str()) == 0)
			continue;

		SecureZeroMemory(&neworient, sizeof(neworient));
		if (!ksCalculateW(const_cast<wchar_t*>(row[FIELD_ANGLEX].c_str()), &neworient.angleX))
			continue;
		if (!ksCalculateW(const_cast<wchar_t*>(row[FIELD_ANGLEY].c_str()), &neworient.angleY))
			continue;
		if (!ksCalculateW(const_cast<wchar_t*>(row[FIELD_ANGLEZ].c_str()), &neworient.angleZ))
			continue;
		wcsncpy(neworient.nameshort, row[FIELD_NAME].c_str(), _countof(neworient.nameshort) - 1);
		if (fields > 4)
			wcsncpy(neworient.namefull, row[FIELD_TOOLTIP].c_str(), _countof(neworient.namefull) - 1);

		m_orients.push_back(neworient);
		count++;
	}

	return (count > 0);
}

static void VersionInfo()
{
	CString ver = TEXT(VER_PRODUCT_NAME);
	ver += L"\n";
	ver += TEXT(VER_FILE_DESCRIPTIONL);
	ver += L"\nВерсия ";
	ver += VER_VERSION_DISPLAY_FULL;
	ver += L"\n";
	ver += TEXT(VER_DATE_AUTHOR);
	MessageBox((HWND)GetHWindow(), (LPCWSTR)ver, L"О библиотеке", MB_OK|MB_ICONINFORMATION);
}

//exports
//-------------------------------------------------------------------------------
unsigned int WINAPI LibraryBmpBeginID (unsigned int bmpSizeType)
{
	switch(bmpSizeType) {
	case ksBmp1616: return 1000;
	case ksBmp2424: return 2000;
	case ksBmp3232: return 3000;
	case ksBmp4848: return 4000;
 	}
	return 0;
}

// Задать идентификатор ресурсов
// ---
unsigned int WINAPI LIBRARYID()
{
  return IDR_LIBID;
}

void WINAPI LIBRARYENTRY( unsigned int comm )
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (comm == IDR_VERINFO)
		return VersionInfo();

	ksAPI7::IKompasDocument3DPtr doc3d( theApp.GetCurrentDocument() );

	if ( doc3d ) {
		theApp.EnableTaskAccess(FALSE);
		switch ( comm ) {
		case IDR_FCSEL: theApp.SelectDialog( doc3d ); break;
		case IDR_FCSET: theApp.ColorChangeDialog( doc3d ); break;
		case IDR_ORIENT: theApp.PlacementDialog( doc3d ); break;
		case IDR_VAREXP: theApp.VariablesExport( doc3d, Document::EMBODIMENT_CURRENT ); break;
		case IDR_VARIMP: theApp.VariablesImport( doc3d, Document::EMBODIMENT_CURRENT ); break;
		}
		theApp.EnableTaskAccess(TRUE);
	} else if (theApp.hasAPI())
		MessageBox((HWND)GetHWindow(), theApp.LoadStr(IDS_NODOC), theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONEXCLAMATION);
}
