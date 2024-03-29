////////////////////////////////////////////////////////////////////////////////
//
// BatchExport.cpp
// BatchExport library
// Alex Fedorov, 2023
//
// Uses function matches() from
// iproute2, licensed under GNU GPL v2
//
// Uses functions GetHexDigit() and UnSlash() from
// Notepad2, licensed under BSD license
//
// Uses function CreateDirectoryRecursively() kindly provided by
// Cygon @ http://blog.nuclex-games.com/2012/06/how-to-create-directories-recursively-with-win32/
//
// Function StrStatus() uses bit manipulation code kindly provided by
// invaliddata @ https://stackoverflow.com/questions/3142867/finding-bit-positions-in-an-unsigned-32-bit-integer
//
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"

#include "BatchExport.h"
#include "GridViewDlg.h"
#include "..\common\OverwriteDlg\OverwriteDlg.h"
#include "..\common\utils.h"
#include "..\common\CSVWriter.h"
#include "..\common\CSVParser.h"

#include "version.h"

// Field names
const LPCWSTR CBatchExportApp::fieldnames[FIELD_COUNT] = { L"Исполнение", L"Директория", L"Имя", L"Комментарий", L"Переменные", L"Опции", L"Статус" };
// Format file extensions, required for every format except native, index corresponds to format option value
const LPCWSTR CBatchExportApp::formatexts[VAROPT_FORMAT_MAX] = { L".stp", L".stp", L".stl", L".wrl" };
// Options map
typedef struct {
	LPCWSTR name; // option name, must be in LOW case
	UINT val;     // bits to set
	UINT mask;    // bits to reset
} OPTIONMAP;
static OPTIONMAP const optionmap[] = {
	// no* options should be at the end, format options should be at the beginning
	{ L"native", CBatchExportApp::VAROPT_FORMAT_NATIVE, CBatchExportApp::VAROPT_FORMAT_MASK },
	{ L"ap203", CBatchExportApp::VAROPT_FORMAT_AP203, CBatchExportApp::VAROPT_FORMAT_MASK },
	{ L"ap214", CBatchExportApp::VAROPT_FORMAT_AP214, CBatchExportApp::VAROPT_FORMAT_MASK },
	{ L"stl", CBatchExportApp::VAROPT_FORMAT_STL, CBatchExportApp::VAROPT_FORMAT_MASK },
	{ L"vrml", CBatchExportApp::VAROPT_FORMAT_VRML, CBatchExportApp::VAROPT_FORMAT_MASK },
	{ L"curves", CBatchExportApp::VAROPT_CURVES, CBatchExportApp::VAROPT_CURVES },
	{ L"dccomm", CBatchExportApp::VAROPT_DELAYED_COMPOSE_COMMENTADD, CBatchExportApp::VAROPT_DELAYED_COMPOSE_COMMENTADD },
	{ L"dcdir", CBatchExportApp::VAROPT_DELAYED_COMPOSE_DIR, CBatchExportApp::VAROPT_DELAYED_COMPOSE_DIR },
	{ L"dcname", CBatchExportApp::VAROPT_DELAYED_COMPOSE_NAME, CBatchExportApp::VAROPT_DELAYED_COMPOSE_NAME },
	{ L"included", CBatchExportApp::VAROPT_INCLUDE, CBatchExportApp::VAROPT_INCLUDE },
	{ L"excluded", 0, CBatchExportApp::VAROPT_INCLUDE },
	{ L"nocurves", 0, CBatchExportApp::VAROPT_CURVES },
	{ L"nodccomm", 0, CBatchExportApp::VAROPT_DELAYED_COMPOSE_COMMENTADD },
	{ L"nodcdir",  0, CBatchExportApp::VAROPT_DELAYED_COMPOSE_DIR },
	{ L"nodcname", 0, CBatchExportApp::VAROPT_DELAYED_COMPOSE_NAME },
	// must be the last
	{ NULL, 0, 0 }
};

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;

typedef enum {
	TYPE_PTR,
	TYPE_LONG,
	TYPE_DOUBLE,
	TYPE_STR
} VALTYPE;

#define showval(type, val) showVAL(L#val, type, val)
void showVAL( LPCWSTR caption, VALTYPE type, ... )
{
	va_list args;
	CString str;

	va_start(args, type);
	switch (type) {
	case TYPE_PTR:
		str.FormatV(L"%p", args);
		break;
	case TYPE_LONG:
		str.FormatV(L"%d", args);
		break;
	case TYPE_DOUBLE:
		str.FormatV(L"%f", args);
		break;
	case TYPE_STR:
		str.FormatV(L"%s", args);
		break;
	default:
		va_end(args);
		return;
	}
	va_end(args);

	MessageBox(NULL, (LPCWSTR)str, caption, MB_OK);
}
#endif /* _DEBUG */

CBatchExportApp theApp;

CBatchExportApp::CBatchExportApp()
{
	m_newKompasAPI = NULL;
	m_kompas = NULL;
	m_overwritefiles = BST_INDETERMINATE;
	m_testmode = FALSE;

	m_numlocale = _wcreate_locale(LC_NUMERIC, L"C");
	m_timelocale = _wcreate_locale(LC_TIME, L"");
}

CBatchExportApp::~CBatchExportApp(){
	_free_locale(m_timelocale);
	_free_locale(m_numlocale);
}

BOOL CBatchExportApp::InitInstance()
{
	// InitCommonControls() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	InitCommonControls();

	AfxEnableControlContainer();

	return CWinApp::InitInstance();
}

BOOL CBatchExportApp::ExitInstance()
{
	m_newKompasAPI = NULL;
	m_kompas = NULL;

	return CWinApp::ExitInstance();
}

// public
ksAPI7::IKompasDocument3DPtr CBatchExportApp::GetCurrentDocument()
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


UINT CBatchExportApp::PreparePath( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATION * variation, CString & pathname )
{
	LPCWSTR ext = L"";

	if (variation->options & CBatchExportApp::VAROPT_FORMAT_MASK) { // export format
		UINT extn = (variation->options & CBatchExportApp::VAROPT_FORMAT_MASK) - 1;
		if (extn >= _countof(formatexts))
			return LIBSTATUS_WARN_FORMAT_UNSUP;
		ext = formatexts[extn];
	} else {														// native format
		ksAPI7::IKompasDocumentPtr doc(doc3d);
		if (doc == NULL)
			return LIBSTATUS_ERR_API;

		switch(doc->DocumentType) {
		case ksDocumentPart:     ext = L".m3d"; break;
		case ksDocumentAssembly: ext = L".a3d"; break;
		default:
			return LIBSTATUS_WARN_FORMAT_UNSUP;
		}
	}

	pathname.Empty();
	if (!PathIsUNC(variation->dir) && !PathIsRoot(variation->dir))
		pathname = L"\\\\?\\";	// convert to UNC to support long paths
	pathname += variation->dir;
	pathname.TrimRight('\\');
	if (!m_testmode && !Utils::CreateDirectoryRecursively(pathname))
		return GetLastError() | LIBSTATUS_SYSERR;

	CString filename = variation->name;
	Utils::SanitizeFileSystemString(filename, '_', TRUE);
	variation->name = filename + ext;
	pathname += '\\' + variation->name;

	if (m_testmode)
		return LIBSTATUS_SUCCESS;

	DWORD fileAttributes = GetFileAttributes((LPCWSTR)pathname);
	if (fileAttributes != INVALID_FILE_ATTRIBUTES) {	// file exists
		if (((fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) ||
			((fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0)) {	// directory is on the way
			return ERROR_CANNOT_MAKE | LIBSTATUS_SYSERR;
		}
		if (m_overwritefiles == BST_INDETERMINATE) {
			OverwriteDlg Dlg(L"BatchExport", OverwriteDlg::ENTITY_FILE, (LPCWSTR)pathname);
			INT_PTR nRet = Dlg.DoModal();
			switch(nRet) {
			case IDNONE:
				m_overwritefiles = BST_UNCHECKED;
			case IDNO:
				return ERROR_FILE_EXISTS | LIBSTATUS_SYSERR;
			case IDALL:
				m_overwritefiles = BST_CHECKED;
			case IDYES:
				break;
			default:
				return LIBSTATUS_ERR_ABORTED;
			}
		} else {
			if (m_overwritefiles != BST_CHECKED) {
				return ERROR_FILE_EXISTS | LIBSTATUS_SYSERR;
			}
		}
	}

	return LIBSTATUS_SUCCESS;
}

UINT CBatchExportApp::SaveAs( ksAPI7::IKompasDocument3DPtr & doc3d, CString & pathname, const VARIATION * variation )
{
	if (m_testmode)
		return LIBSTATUS_INFO_TESTOK;

	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return LIBSTATUS_ERR_API;

	BOOL ret = doc3d5->SaveAs(pathname.LockBuffer());
	pathname.UnlockBuffer();
	if (!ret) {
		DWORD err = GetLastError();
		if (err != ERROR_SUCCESS)
			return (UINT)err | LIBSTATUS_SYSERR;
		return LIBSTATUS_ERR_API;
	}
	return LIBSTATUS_SUCCESS;
}

UINT CBatchExportApp::ExportFile( ksAPI7::IKompasDocument3DPtr & doc3d, CString & pathname, const VARIATION * variation )
{
	if (m_testmode)
		return LIBSTATUS_INFO_TESTOK;

	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return LIBSTATUS_ERR_API;

	IAdditionFormatParamPtr param = doc3d5->AdditionFormatParam();
	if (param == NULL)
		return LIBSTATUS_ERR_API;

	UINT options = variation->options;
	param->Init();
	switch(options & VAROPT_FORMAT_MASK) {
	case VAROPT_FORMAT_AP203:
		param->SetFormat(format_STEP);				// export to AP203, default system behavior
		param->SetFormatBinary(FALSE);
		param->SetTopolgyIncluded(FALSE);
		break;
	case VAROPT_FORMAT_AP214:						// extended options that "not included out of the box"
		param->SetFormat(format_STEP);
		param->SetFormatBinary(TRUE);				// export to AP214
		if (options & VAROPT_CURVES)
			param->SetTopolgyIncluded(TRUE);
		else
			param->SetTopolgyIncluded(FALSE);
		break;
	case VAROPT_FORMAT_STL:
		param->SetFormat(format_STL);
		param->SetFormatBinary(TRUE);				// binary
		param->SetTopolgyIncluded(FALSE);			// export surfaces
		break;
	case VAROPT_FORMAT_VRML:
		param->SetFormat(format_VRML);
		param->SetFormatBinary(FALSE);				// N/A
		param->SetTopolgyIncluded(FALSE);			// N/A
		break;
	default:
		return LIBSTATUS_WARN_FORMAT_UNSUP;
	}

	BOOL ret = doc3d5->SaveAsToAdditionFormat(pathname.LockBuffer(), param);
	pathname.UnlockBuffer();
	if (!ret) {
		DWORD err = GetLastError();
		if (err != ERROR_SUCCESS)
			return (UINT)err | LIBSTATUS_SYSERR;
		return LIBSTATUS_ERR_API;
	}
	return LIBSTATUS_SUCCESS;
}

BOOL CBatchExportApp::BatchExport( ksAPI7::IKompasDocument3DPtr & doc3d )
{
	VARIATIONS variations;
	UINT ret = LIBSTATUS_SUCCESS;
	UINT processed = 0, exported = 0;
	BOOL statdlg = FALSE;

	// reset state
	m_overwritefiles = BST_INDETERMINATE;
	m_testmode = FALSE;
	m_quick = FALSE;

	if (GetKeyState(VK_CONTROL) & 0x80) { m_quick = TRUE; m_overwritefiles = BST_CHECKED; }
	else if (GetKeyState(VK_MENU) & 0x80) m_overwritefiles = BST_UNCHECKED;
	else if (GetKeyState(VK_SHIFT) & 0x80) m_overwritefiles = BST_CHECKED;

	if ((ret = VariationsEnum( doc3d, &variations)) == LIBSTATUS_SUCCESS) {
		size_t size = variations.size();
		size_t total = 0;
		VARIATION *varptr = NULL;

		for(INT i = 0; i < (INT)size; i++) {
			if ((variations[i].options & VAROPT_INCLUDE) == VAROPT_INCLUDE)
				total++;
		}
		if (total == 0) {
			ret = LIBSTATUS_ERR_VARIATIONS_EXCLUDED;
			goto exit;
		}

		// save current embodiment
		LONG orige;
		if (!Document::EmbodimentGet(doc3d, orige)) {
			ret = LIBSTATUS_ERR_API;
			goto exit;
		}
		// save comment
		LPCWSTR cstrptr = NULL;
		if (!Document::CommentGet(doc3d, cstrptr)) {
			ret = LIBSTATUS_ERR_API;
			goto exit;
		}
		CString origcomment(cstrptr);

		CString pstr;
		ksAPI7::IProgressBarIndicatorPtr progress(m_newKompasAPI->GetProgressBarIndicator());
		if (progress != NULL)
			progress->Start(0, (LONG)total, L"Экспорт...", TRUE);

		for(INT i = 0; i < (INT)size; i++) {
			varptr = &variations[i];
			if ((varptr->options & VAROPT_INCLUDE) != VAROPT_INCLUDE) {
				continue;
			}
			if (ret != LIBSTATUS_SUCCESS) {
				varptr->status = LIBSTATUS_ERR_DISCARDED;
				continue;
			}

			if (!Document::EmbodimentSet(doc3d, varptr->embodiment)) {
				varptr->status = ret = LIBSTATUS_ERR_SET_EMBODIMENT;
				continue;
			}

			size_t varsz = varptr->variables.size();
			for(INT v = 0; v < (INT)varsz; v++) {
				VARIABLE *vptr = &varptr->variables[v];
				if (!VariableSet(doc3d, vptr->name, vptr, vptr->opt & VOPT_LOCAL?Document::EMBODIMENT_CURRENT:Document::EMBODIMENT_TOP)) {
					varptr->status = ret = LIBSTATUS_ERR_SET_VARIABLE;
					break;
				}
			}
			if (ret != LIBSTATUS_SUCCESS)
				continue;

			if (!Document::RebuildDocument(doc3d)) {
				varptr->status = ret = LIBSTATUS_ERR_DOCUMENT_REBUILD;
				continue;
			} 

			if (!Document::IsDocumentValid(doc3d)) {
				varptr->status = ret = LIBSTATUS_ERR_DOCUMENT_INVALID;
				continue;
			}

			// Delayed composition
			// only update time when necessary
			if ((varptr->options & VAROPT_DELAYED_COMPOSE_MASK) &&
				!VariationUpdateTime(varptr)) {
				varptr->status = ret = LIBSTATUS_ERR_CURRENT_TIME;
				continue;
			}
			if (varptr->options & VAROPT_DELAYED_COMPOSE_DIR) {
				CString templ(varptr->dir);
				varptr->status = ComposeDir( doc3d, templ, varptr );
				if (varptr->status != LIBSTATUS_SUCCESS) {
					ret = varptr->status;
					continue;
				}
			}
			if (varptr->options & VAROPT_DELAYED_COMPOSE_NAME) {
				CString templ(varptr->name);
				varptr->status = ComposeString( doc3d, templ, varptr, VP_NAME );
				if (varptr->status != LIBSTATUS_SUCCESS) {
					ret = varptr->status;
					continue;
				}
			}
			if (varptr->options & VAROPT_DELAYED_COMPOSE_COMMENTADD) {
				CString templ(varptr->commentadd);
				varptr->status = ComposeString( doc3d, templ, varptr, VP_COMMENTADD );
				if (varptr->status != LIBSTATUS_SUCCESS) {
					ret = varptr->status;
					continue;
				}
			}

			LPWSTR strptr;
			if (!Document::ModelNameGet(doc3d, strptr)) {
				varptr->status = ret = LIBSTATUS_ERR_API;
				continue;
			}
			CString origname(strptr);
			BOOL namerq = varptr->name.IsEmpty()?FALSE:TRUE;
			if (namerq) {
				if (!Document::ModelNameSet(doc3d, varptr->name)) {
					varptr->status = ret = LIBSTATUS_ERR_SET_NAME;
					continue;
				}
			} else
				varptr->name = origname;

			if (!varptr->commentadd.IsEmpty()) {
				CString comment(origcomment);

				comment.Append(varptr->commentadd);
				// Kompass will crash on handling comments larger than 511 chars
				if (comment.GetLength() > 511)
					comment.Truncate(511);
				LPWSTR comptr = comment.LockBuffer();
				Utils::UnSlash(comptr);
				Utils::SanitizeString(comptr, L"*", L'_');
				BOOL lret = Document::CommentSet(doc3d, comptr);
				comment.UnlockBuffer();
				if (!lret) {
					varptr->status = ret = LIBSTATUS_ERR_SET_COMMENT;
					continue;
				}
			} else {
				// restore comment
				LPWSTR comptr = origcomment.LockBuffer();
				Utils::SanitizeString(comptr, L"*", L'_');
				BOOL lret = Document::CommentSet(doc3d, comptr);
				origcomment.UnlockBuffer();
				if (!lret) {
					varptr->status = ret = LIBSTATUS_ERR_SET_COMMENT;
					continue;
				}
			}

			processed++;
			if (progress != NULL) {
				pstr.Format(L"%s (%d/%d)", varptr->name, processed, total);
				progress->SetProgress(processed, (LPCWSTR)pstr, TRUE);
			}

			CString pathname;
			if ((varptr->status = PreparePath(doc3d, varptr, pathname)) != LIBSTATUS_SUCCESS) {
				if ((varptr->status & LIBSTATUS_EXCLUSIVE) == LIBSTATUS_ERR)
					ret = varptr->status; // abort operation on fatal export errors (excluding system errors)
				continue;
			}

			if (varptr->options & VAROPT_FORMAT_MASK)	// export format
				varptr->status = ExportFile(doc3d, pathname, varptr);
			else										// native format
				varptr->status = SaveAs(doc3d, pathname, varptr);

			if ((varptr->status & LIBSTATUS_EXCLUSIVE) == LIBSTATUS_ERR)
				ret = varptr->status; // abort operation on fatal export errors

			// restore original embodiment name
			if (namerq && !Document::ModelNameSet(doc3d, origname))
				varptr->status = ret = LIBSTATUS_ERR_SET_NAME;

			if (varptr->status == LIBSTATUS_SUCCESS) { // variation is done if no error occured so far
				varptr->status = LIBSTATUS_INFO_DONE;
				exported++;
			}
		}

		if (progress != NULL)
			progress->Stop(L"Готово", TRUE);

		// don't restore original embodiment
		/*if (!m_testmode)
			EmbodimentSet(doc3d, orige);*/
		// restore comment
		Document::CommentSet(doc3d, origcomment.LockBuffer());
		origcomment.UnlockBuffer();

		// show status
		if ( !(m_quick && exported == total) &&
			 !(m_testmode && total == 1) ) {
			GridViewDlg *Dlg = new GridViewDlg(L"Отчет", L"Закрыть", doc3d, &variations,
				GridViewDlg::OPTIONS(GridViewDlg::GVOPT_READONLY
									|GridViewDlg::GVOPT_SKIPEXCL
									|GridViewDlg::GVOPT_SHOWVARIABLES
									|GridViewDlg::GVOPT_SHOWSTATUS
									|GridViewDlg::GVOPT_ERRIGNORE
									));
			Dlg->DoModal();
			delete Dlg;
			statdlg = TRUE;
		}

	}

exit:
	if (!statdlg && (ret != LIBSTATUS_SUCCESS)) {
		if (ret == LIBSTATUS_ERR_ABORTED)
			return FALSE;
		CString errstr;
		StrStatus(errstr, ret);
		MessageBox((HWND)GetHWindow(), errstr, LoadStr(IDR_LIBID), MB_OK|MB_ICONEXCLAMATION);
	}
	return (ret == LIBSTATUS_SUCCESS);
}



BOOL CBatchExportApp::VariableSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, VARIABLE * var, LONG embodiment )
{
	if (var->curval >= var->values.size())
		return FALSE;
	if (!var->values[var->curval].expr.IsEmpty())
		return Document::VariableExprSet(doc3d, varname, var->values[var->curval].expr, embodiment);
	else
		return Document::VariableValueSet(doc3d, varname, var->values[var->curval].value, embodiment);

	return FALSE;
}



LONG CBatchExportApp::VariationsGetRowCount( reference pAttr )
{
	if (pAttr == NULL)
		return -1;

	ksAttributeObjectPtr attrObj( m_kompas->GetAttributeObject() );
	if (attrObj == NULL)
		return -1;

	LONG rowcount = 0, colcount = 0;
	if (!attrObj->ksGetAttrTabInfo(pAttr, &rowcount, &colcount))
		return -1;

	return rowcount;
}

BOOL CBatchExportApp::VariationGet( reference pAttr, LONG row, ksDynamicArrayPtr & valarray )
{
	if (pAttr == 0)
		return FALSE;

	ksAttributeObjectPtr attrObj( m_kompas->GetAttributeObject() );
	if (attrObj == NULL)
		return FALSE;

	ksUserParamPtr values( m_kompas->GetParamStruct( ko_UserParam ) );
	if (values == NULL)
		return FALSE;
	ksDynamicArrayPtr newarray( m_kompas->GetDynamicArray( LTVARIANT_ARR ) );
	if (newarray == NULL)
		return FALSE;
	ksLtVariantPtr val( m_kompas->GetParamStruct( ko_LtVariant ) );
	val->Init();
	val->intVal = 0;
	newarray->ksAddArrayItem( -1, val ); // embodiment
	val->wstrVal = L"";
	newarray->ksAddArrayItem( -1, val ); // dir
	newarray->ksAddArrayItem( -1, val ); // name
	newarray->ksAddArrayItem( -1, val ); // commentadd
	newarray->ksAddArrayItem( -1, val ); // variables
	newarray->ksAddArrayItem( -1, val ); // options
	values->Init();
	values->SetUserArray( newarray );

	if (!attrObj->ksGetAttrRow( pAttr, row, NULL, NULL, values )) {
		newarray->ksDeleteArray();
		return FALSE;
	}

	if (valarray != NULL)
		valarray->ksDeleteArray();
	valarray = newarray;

	return TRUE;
}

BOOL CBatchExportApp::VariationSet( reference pAttr, LONG row, ksDynamicArrayPtr & valarray )
{
	if (pAttr == 0)
		return FALSE;

	ksAttributeObjectPtr attrObj( m_kompas->GetAttributeObject() );
	if (attrObj == NULL)
		return FALSE;

	LONG rowcount = 0, colcount = 0;
	if (!attrObj->ksGetAttrTabInfo(pAttr, &rowcount, &colcount))
		return FALSE;

	ksUserParamPtr values( m_kompas->GetParamStruct( ko_UserParam ) );
	if (values == NULL)
		return FALSE;
	values->Init();
	values->SetUserArray( valarray );

	if (row == -1 || row >= rowcount)
		return attrObj->ksAddAttrRow( pAttr, -1, NULL, values, L"" );

	return attrObj->ksSetAttrRow( pAttr, row, NULL, NULL, values, L"" );
}

BOOL CBatchExportApp::VariationUpdateTime( tm * ctime )
{
	time_t curtime = time(NULL);
	if (curtime == (time_t)-1)
		return FALSE;
	if (localtime_s(ctime, &curtime) != 0)
		return FALSE;
	return TRUE;
}

BOOL CBatchExportApp::VariationUpdateTime( VARIATION * variation )
{
	return VariationUpdateTime(&variation->time);
}

reference CBatchExportApp::VariationsCreateAttr( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR lib )
{
	ksAttributeT parAttribute; // Структура параметров табличного атрибута
	memset( &parAttribute, 0, sizeof(parAttribute) );
	parAttribute.key1 = 3;
	parAttribute.key2 = 0;
	parAttribute.key3 = 0;
	parAttribute.key4 = 1100; // created by the lib

	return ksCreateAttrT( doc3d->GetReference(), &parAttribute, VARTYPENUM, lib);
}

#define SETITEM( item ) \
			if (!values->ksSetArrayItem( item, val )) { \
				ret = LIBSTATUS_ERR_API; \
				break; \
			}
UINT CBatchExportApp::VariationsSave( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations )
{
	UINT ret = LIBSTATUS_SUCCESS;

	size_t count = variations->size();

	ksDynamicArrayPtr values( m_kompas->GetDynamicArray( LTVARIANT_ARR ) );
	if (values == NULL)
		return LIBSTATUS_ERR_API;
	ksLtVariantPtr val( m_kompas->GetParamStruct( ko_LtVariant ) );
	if (val == NULL) {
		values->ksDeleteArray();
		return LIBSTATUS_ERR_API;
	}

	val->Init();
	val->intVal = 0;
	values->ksAddArrayItem( -1, val ); // embodiment
	val->wstrVal = L"";
	values->ksAddArrayItem( -1, val ); // dir
	values->ksAddArrayItem( -1, val ); // name
	values->ksAddArrayItem( -1, val ); // commentadd
	values->ksAddArrayItem( -1, val ); // variables
	values->ksAddArrayItem( -1, val ); // options

	for (size_t i = 0; i < count; i++) {
		VARIATION *varptr = &variations->at(i);
		if (varptr->AttrPtr == NULL)
			continue;
		if (varptr->AttrRow > VariationsGetRowCount(varptr->AttrPtr))
			continue;
		val->intVal = varptr->embodiment;
		SETITEM(VP_EMBODIMENT);
		val->wstrVal = (LPCWSTR)varptr->dir;
		SETITEM(VP_DIR);
		val->wstrVal = (LPCWSTR)varptr->name;
		SETITEM(VP_NAME);
		val->wstrVal = (LPCWSTR)varptr->commentadd;
		SETITEM(VP_COMMENTADD);
		val->wstrVal = (LPCWSTR)varptr->varstr;
		SETITEM(VP_VARIABLES);
		val->wstrVal = (LPCWSTR)varptr->optstr;
		SETITEM(VP_OPTIONS);

		if (!VariationSet(varptr->AttrPtr, varptr->AttrRow, values)) {
			ret = LIBSTATUS_ERR_API;
			break;
		}
	}
	values->ksDeleteArray();

	return ret;
}

#define GETITEM( item ) \
			if (!values->ksGetArrayItem( item, val )) { \
				ret = LIBSTATUS_ERR_API; \
				break; \
			}
UINT CBatchExportApp::VariationsGetTemplates( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * templates )
{
	UINT ret = LIBSTATUS_SUCCESS;
	tm vartime = {};

	LONG emcount = Document::EmbodimentCount(doc3d);
	if (emcount <= 0)
		return LIBSTATUS_ERR_NO_EMBODIMENTS;

	// init time
	if (!VariationUpdateTime(&vartime))
		return LIBSTATUS_ERR_CURRENT_TIME;

	ksDynamicArrayPtr values( NULL );
	reference rAttr = CreateAttrIterator( doc3d->GetReference(), 0, 0, 0, 0, VARTYPENUM );
	if (rAttr == 0)
		return LIBSTATUS_ERR_API;

	LONG count = 0;
	reference pAttr = MoveAttrIterator( rAttr, 'F', NULL );
	if (pAttr == NULL) {
		DeleteIterator( rAttr );
		return LIBSTATUS_ERR_NO_ATTRS;
	}
	for(; pAttr != NULL; pAttr = MoveAttrIterator( rAttr, 'N', NULL )) {
		LONG rows = VariationsGetRowCount(pAttr);
		for(LONG row = 0; row < rows; row++) {
			count++;
			if (!VariationGet(pAttr, row, values)) {
				DeleteIterator( rAttr );
				return LIBSTATUS_ERR_API;
			}
			templates->push_back(VARIATION());
			VARIATION* varptr = &templates->back();
			if (varptr == NULL) {
				DeleteIterator( rAttr );
				return LIBSTATUS_SYSERR_NO_MEM;
			}

			varptr->AttrPtr = pAttr;
			varptr->AttrRow = row;
			varptr->options = VAROPT_DEFAULT;
			varptr->status = LIBSTATUS_SUCCESS;
			ksLtVariantPtr val( m_kompas->GetParamStruct( ko_LtVariant ) );
			if (val == NULL) {
				values->ksDeleteArray();
				DeleteIterator( rAttr );
				return LIBSTATUS_ERR_API;
			}

			GETITEM(VP_EMBODIMENT);
			varptr->embodiment = val->longVal;
			if (varptr->embodiment >= emcount)
				varptr->status |= LIBSTATUS_BAD_EMBODIMENT;
			GETITEM(VP_DIR);
			varptr->status |= ComposeDir(doc3d, val->wstrVal, varptr);
			varptr->dir.SetString(val->wstrVal); // restore as template
			GETITEM(VP_NAME);
			varptr->status |= ComposeString(doc3d, val->wstrVal, varptr, VP_NAME);
			varptr->name.SetString(val->wstrVal);
			GETITEM(VP_COMMENTADD);
			varptr->status |= ComposeString(doc3d, val->wstrVal, varptr, VP_COMMENTADD);
			varptr->commentadd.SetString(val->wstrVal);
			GETITEM(VP_VARIABLES);
			varptr->varstr.SetString(val->wstrVal);
			varptr->status |= ParseVariables(doc3d, varptr->varstr, varptr);
			GETITEM(VP_OPTIONS);
			varptr->optstr.SetString(val->wstrVal);
			varptr->status |= ParseOptions(varptr->optstr, &varptr->options);

			if (varptr->status != LIBSTATUS_SUCCESS)
				m_quick = FALSE;
		}
	}
	values->ksDeleteArray();
	DeleteIterator( rAttr );

	return ret;
}

UINT CBatchExportApp::VariationsEnum( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations )
{
	VARIATIONS templates;

	UINT ret = VariationsGetTemplates( doc3d, &templates);

	if (ret != LIBSTATUS_SUCCESS)
		return ret;
	if (templates.size() < 1)
		return LIBSTATUS_ERR_NO_VARIATIONS;

	BOOL repeat = FALSE;
	do {
		INT_PTR nRet = IDOK;
		if (!m_quick) {
			GridViewDlg *Dlg = new GridViewDlg(L"Список шаблонов", L"Далее >", doc3d, &templates,
				GridViewDlg::OPTIONS(GridViewDlg::GVOPT_PRESERVETEMPL
									|GridViewDlg::GVOPT_ALLOWSAVE
									|GridViewDlg::GVOPT_SHOWCANCEL
									|GridViewDlg::GVOPT_SHOWINCLUDE
									|GridViewDlg::GVOPT_SHOWVARIABLES
									));
			nRet = Dlg->DoModal();
			delete Dlg;
		}
		if (nRet == IDOK) {
			size_t varattrcount = templates.size();
			UINT processed = 0;
			variations->clear();
			for(int i = 0; i < (int)varattrcount; i++) {
				VARIATION *varptr = &templates[i];
				if (varptr->options & VAROPT_INCLUDE) {
					processed++;
					ret |= DoRange(doc3d, variations, varptr, 0);	// DoRange only set non exclusive input errors
				}
			}
			if (processed == 0)
				return LIBSTATUS_ERR_VARIATIONS_EXCLUDED;
			size_t varcount = variations->size();
			if (ret != LIBSTATUS_SUCCESS) {
				CString errstr;
				StrStatus(errstr, ret);
				if (ret > LIBSTATUS_USER_ERR_MAX)
					return ret;
				// should never happen due to previous input error checking, but who knows...
				errstr.Insert(0, L"Некоторые вариации могли быть исключены в связи со следующими ошибами:\n");
				MessageBox((HWND)GetHWindow(), (LPCWSTR)errstr, LoadStr(IDR_LIBID), MB_OK|MB_ICONEXCLAMATION);
				m_quick = FALSE;
			}
			if (varcount < 1)
				return LIBSTATUS_ERR_NO_VARIATIONS;
			if (m_quick)
				return LIBSTATUS_SUCCESS;

			GridViewDlg *Dlg = new GridViewDlg(L"Список вариаций", L"Экспорт", doc3d, variations,
				GridViewDlg::OPTIONS(GridViewDlg::GVOPT_READONLY
									|GridViewDlg::GVOPT_CANBACK
									|GridViewDlg::GVOPT_SHOWCANCEL
									|GridViewDlg::GVOPT_SHOWFILEOPTS
									|GridViewDlg::GVOPT_SHOWINCLUDE
									|GridViewDlg::GVOPT_SHOWVARIABLES
									));
			Dlg->SetOverWriteFiles(m_overwritefiles);
			Dlg->SetTestMode(m_testmode);
			nRet = Dlg->DoModal();
			if ((nRet == IDOK) || (nRet == IDC_BACK)) {
				m_overwritefiles = Dlg->GetOverwriteFiles();
				m_testmode = Dlg->GetTestMode();
				delete Dlg;
				if (nRet == IDC_BACK) {
					repeat = TRUE;
					continue;
				}
				return LIBSTATUS_SUCCESS;
			}
			delete Dlg;
			break;
		}
		break;
    } while(repeat); // do attributes

	return LIBSTATUS_ERR_ABORTED;
}

UINT CBatchExportApp::VariationsExport( ksAPI7::IKompasDocument3DPtr & doc3d, BOOL file )
{
	UINT ret = LIBSTATUS_SUCCESS;

	ksDynamicArrayPtr values( NULL );
	reference rAttr = CreateAttrIterator( doc3d->GetReference(), 0, 0, 0, 0, VARTYPENUM );
	if (rAttr == 0)
		return LIBSTATUS_ERR_API;

	reference pAttr = MoveAttrIterator( rAttr, 'F', NULL );
	if (pAttr == NULL)
		return LIBSTATUS_ERR_NO_ATTRS;

	OPENFILENAME *filename = NULL;
	ret = Utils::GetFileName( TRUE, (LPCWSTR)doc3d->GetPathName(), L"Экспорт шаблонов", L"_variations", Utils::CSVfilter, Utils::CSVext, filename);
	if (ret != LIBSTATUS_SUCCESS) {
		DeleteIterator( rAttr ), rAttr = 0;
		if (ret == LIBSTATUS_ERR_ABORTED)
			return LIBSTATUS_SUCCESS;
		return ret;
	}

	CSVWriter csv(FIELD_COUNT - 1); // - VP_STATUS
	for(INT i = 0; i < FIELD_COUNT - 1; i++)
		csv << fieldnames[i];
	csv.newRow();
	LONG attrcount = 0;
	LONG varcount = 0;
	for(; pAttr != NULL; pAttr = MoveAttrIterator( rAttr, 'N', NULL )) {
		LONG rows = VariationsGetRowCount(pAttr);
		if (rows <= 0)
			continue;
		if (attrcount++) {
			csv.newRow(); // attribute separator
			csv.newRow();
		}
		for(LONG row = 0; row < rows; row++) {
			varcount++;
			if (!VariationGet(pAttr, row, values))
				return LIBSTATUS_ERR_API;

			ksLtVariantPtr val( m_kompas->GetParamStruct( ko_LtVariant ) );
			if (val == NULL) {
				values->ksDeleteArray();
				return LIBSTATUS_ERR_API;
			}

			GETITEM(VP_EMBODIMENT);
			csv << val->longVal;
			GETITEM(VP_DIR);
			csv << val->wstrVal;
			GETITEM(VP_NAME);
			csv << val->wstrVal;
			GETITEM(VP_COMMENTADD);
			csv << val->wstrVal;
			GETITEM(VP_VARIABLES);
			csv << val->wstrVal;
			GETITEM(VP_OPTIONS);
			csv << val->wstrVal;
		}
	}
	values->ksDeleteArray();
	DeleteIterator(rAttr);

	if (varcount < 1)
		return LIBSTATUS_ERR_NO_VARIATIONS;

	SetLastError(ERROR_SUCCESS);
	if (!csv.writeToFile(filename->lpstrFile)) {
		DWORD err = GetLastError();
		if (err != ERROR_SUCCESS)
			ret = LIBSTATUS_SYSERR | err;
		else
			ret = LIBSTATUS_ERR_UNKNOWN;
	}

	Utils::FreeFileName(filename);

	if (varcount > 0) {
		CString message;
		message.Format(L"Обработано атрибутов: %d\nЭкспортировано шаблонов: %d", attrcount, varcount);
		MessageBox((HWND)GetHWindow(), message, theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONINFORMATION);
	}

	return ret;
}

UINT CBatchExportApp::VariationsImport( ksAPI7::IKompasDocument3DPtr & doc3d, BOOL file)
{
	UINT ret = LIBSTATUS_SUCCESS;
	std::wifstream csvstream;
	aria::csv::CsvParser *parser = NULL;
	VARIATION templ = {};
	VARIATIONS templates;
	LONG AttrNum, AttrRow;
	LONG attrdeleted = 0;
	LONG attradded = 0;
	LONG varcount = 0;
	OPENFILENAME *filename = NULL;
	LPWSTR lib = NULL;

	LONG emcount = Document::EmbodimentCount(doc3d);
	if (emcount <= 0)
		return LIBSTATUS_ERR_NO_EMBODIMENTS;

	if (!GetAttributeLib(lib))
		ret = LIBSTATUS_ERR_ABORTED;

	ksLtVariantPtr val( m_kompas->GetParamStruct( ko_LtVariant ) );
	if (val == NULL)
		ret = LIBSTATUS_ERR_API;
	ksDynamicArrayPtr values( m_kompas->GetDynamicArray( LTVARIANT_ARR ) );
	if (values == NULL)
		ret = LIBSTATUS_ERR_API;

	if (ret != LIBSTATUS_SUCCESS)
		goto cleanup;

	ret = Utils::GetFileName( FALSE, (LPCWSTR)doc3d->GetPathName(), L"Импорт шаблонов", L"_variations", Utils::CSVfilter, Utils::CSVext, filename);
	if (ret != LIBSTATUS_SUCCESS) {
		if (ret == LIBSTATUS_ERR_ABORTED)
			ret = LIBSTATUS_SUCCESS;
		goto cleanup;
	}

	if ((ret = Utils::OpenCSVFile(csvstream, filename->lpstrFile)) != ERROR_SUCCESS)
		goto cleanup;
	parser = new aria::csv::CsvParser(csvstream);
	parser->delimiter(L';');
	for (auto& row : *parser) {
		SIZE_T fields = row.size();
		if (fields == 0 && templ.AttrRow > 0) {
			templ.AttrPtr++;
			templ.AttrRow = 0;
			continue;
		}
		if (fields != (FIELD_COUNT - 1)) // - VP_STATUS
			continue;
		templ.options = VAROPT_INCLUDE; // include item by default, all other options are ignored
		templ.status = LIBSTATUS_SUCCESS;
		LPWSTR endptr = NULL;
		templ.embodiment = wcstol(row[VP_EMBODIMENT].c_str(), &endptr, 0);
		if ((endptr == NULL) || (*endptr != '\0') ||
			(templ.embodiment < 0) || (templ.embodiment == LONG_MAX))
			continue;
		templ.dir = row[VP_DIR].c_str();
		templ.name = row[VP_NAME].c_str();
		templ.commentadd = row[VP_COMMENTADD].c_str();
		templ.varstr = row[VP_VARIABLES].c_str();
		templ.optstr = row[VP_OPTIONS].c_str();
		templates.push_back(templ);
		templ.AttrRow++;
	}

	GridViewDlg *Dlg = new GridViewDlg(L"Импорт шаблонов", L"Импорт", doc3d, &templates,
		GridViewDlg::OPTIONS(GridViewDlg::GVOPT_PRESERVETEMPL
							|GridViewDlg::GVOPT_SHOWCANCEL
							|GridViewDlg::GVOPT_SHOWINCLUDE
							|GridViewDlg::GVOPT_SHOWVARIABLES
							|GridViewDlg::GVOPT_ERRIGNORE
							|GridViewDlg::GVOPT_NOCOMPINC
							), GridViewDlg::GROUP_BY_ATTR);
	INT_PTR nRet = Dlg->DoModal();
	if (nRet != IDOK) {
		ret = LIBSTATUS_ERR_ABORTED;
		goto cleanup;
	}

	ULONG templcnt = 0;
	for (auto& templ : templates)
		if (templ.options & VAROPT_INCLUDE)
			templcnt++;
	if (templcnt == 0) // no templates to import
		goto cleanup;

	reference rAttrIter = CreateAttrIterator( doc3d->GetReference(), 0, 0, 0, 0, VARTYPENUM );
	if (rAttrIter == 0) {
		ret = LIBSTATUS_ERR_API;
		goto cleanup;
	}
	reference rAttr = MoveAttrIterator( rAttrIter, 'F', NULL );
	if (rAttr) {
		std::vector<reference> rAttrs;
		do {
			rAttrs.push_back(rAttr);
			rAttr = MoveAttrIterator( rAttrIter, 'N', NULL );
		} while(rAttr);
		CString message;
		message.Format(L"Документ уже содержит %d атрибутов с UID " UID L", заменить?", rAttrs.size());
		switch(MessageBox((HWND)GetHWindow(), message, theApp.LoadStr(IDR_LIBID), MB_YESNOCANCEL|MB_ICONQUESTION)) {
		case IDCANCEL:
			DeleteIterator( rAttrIter );
			ret = LIBSTATUS_ERR_ABORTED;
			goto cleanup;
		case IDYES:
			for(auto &rAttr: rAttrs) {
				DeleteAttrT(doc3d->GetReference(), rAttr, L"");
				attrdeleted++;
			}
		}
	}
	DeleteIterator( rAttrIter ), rAttrIter = 0;

	val->Init();
	val->intVal = 0;
	values->ksAddArrayItem( -1, val ); // embodiment
	val->wstrVal = L"";
	values->ksAddArrayItem( -1, val ); // dir
	values->ksAddArrayItem( -1, val ); // name
	values->ksAddArrayItem( -1, val ); // commentadd
	values->ksAddArrayItem( -1, val ); // variables
	values->ksAddArrayItem( -1, val ); // options

	AttrNum = -1;
	AttrRow = 0;
	for (auto& templ : templates) {
		// skip excluded
		if (!(templ.options & VAROPT_INCLUDE))
			continue;

		// add a new attribute if needed
		if (AttrNum != templ.AttrPtr) {
			AttrNum = templ.AttrPtr;
			if ((rAttr = VariationsCreateAttr( doc3d, lib )) == 0) {
				ret = LIBSTATUS_ERR_API;
				break;
			}
			AttrRow = 0;
			attradded++;
		}

		// add rows
		val->intVal = templ.embodiment;
		SETITEM(VP_EMBODIMENT);
		val->wstrVal = (LPCWSTR)templ.dir;
		SETITEM(VP_DIR);
		val->wstrVal = (LPCWSTR)templ.name;
		SETITEM(VP_NAME);
		val->wstrVal = (LPCWSTR)templ.commentadd;
		SETITEM(VP_COMMENTADD);
		val->wstrVal = (LPCWSTR)templ.varstr;
		SETITEM(VP_VARIABLES);
		val->wstrVal = (LPCWSTR)templ.optstr;
		SETITEM(VP_OPTIONS);

		if (!VariationSet(rAttr, AttrRow++, values)) {
			ret = LIBSTATUS_ERR_API;
			break;
		}

		varcount++;
	}

cleanup:
	if (values != NULL)
		values->ksDeleteArray();
	if (filename != NULL)
		Utils::FreeFileName(filename);
	if (lib != NULL)
		delete[] lib;
	if (parser != NULL)
		delete parser;

	if (attrdeleted > 0 || varcount > 0) {
		CString message;
		message.AppendFormat(L"Удалено атрибутов: %d\nСоздано атрибутов: %d\nИмпортировано шаблонов: %d", attrdeleted, attradded, varcount);
		MessageBox((HWND)GetHWindow(), message, theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONINFORMATION);
	}

	return ret;
}

UINT CBatchExportApp::VariationsAddVariation( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations, VARIATION * variation )
{
	variations->push_back(VARIATION());
	VARIATION *last = &variations->back();
	last->AttrPtr = variation->AttrPtr;
	last->AttrRow = variation->AttrRow;
	last->embodiment = variation->embodiment;
	last->options = variation->options;
	last->optstr = variation->optstr;
	for(size_t vari = 0; vari < variation->variables.size(); vari++) {
		auto *srcvar = &variation->variables.at(vari);
		VARIABLE newvar;
		last->variables.push_back(newvar);
		auto *dstvar = &last->variables.back();
		dstvar->values.push_back(srcvar->values[srcvar->curval]);
		dstvar->name = srcvar->name;
		dstvar->opt = srcvar->opt;
		dstvar->curval = 0;
	}

	// only update time when necessary
	if (((last->options & VAROPT_DELAYED_COMPOSE_MASK) != VAROPT_DELAYED_COMPOSE_MASK) &&
		!VariationUpdateTime(last))
		return LIBSTATUS_ERR_CURRENT_TIME;
	if (last->options & VAROPT_DELAYED_COMPOSE_DIR)
		last->dir = variation->dir;
	else {
		last->status = ComposeDir( doc3d, variation->dir, last);
		if (last->status != LIBSTATUS_SUCCESS)
			return last->status;
	}
	if (last->options & VAROPT_DELAYED_COMPOSE_NAME)
		last->name = variation->name;
	else {
		last->status = ComposeString( doc3d, variation->name, last, VP_NAME);
		if (last->status != LIBSTATUS_SUCCESS)
			return last->status;
	}
	if (last->options & VAROPT_DELAYED_COMPOSE_COMMENTADD)
		last->commentadd = variation->commentadd;
	else {
		last->status = ComposeString( doc3d, variation->commentadd, last, VP_COMMENTADD);
		if (last->status != LIBSTATUS_SUCCESS)
			return last->status;
	}

	last->status = ComposeVariables( doc3d, &last->variables, &last->varstr, TRUE);
	if (last->status != LIBSTATUS_SUCCESS)
		return last->status;

	return LIBSTATUS_SUCCESS;
}

UINT CBatchExportApp::DoRange( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations, VARIATION * variation, ULONG index )
{
	if (index < 0)
		return LIBSTATUS_SUCCESS;
	if (variation->variables.size() == 0)
		return VariationsAddVariation( doc3d, variations, variation );

	UINT ret = 0;
	VARIABLE *var = &variation->variables[index];
	for(ULONG vali = 0; vali < (ULONG)var->values.size(); vali++) {
		var->curval = vali;
		VARVALUE *val = &var->values[vali];
		if (!val->expr.IsEmpty()) {
			if (index < (variation->variables.size() - 1)) {
				INT lret = DoRange(doc3d, variations, variation, index + 1);
				if (lret != LIBSTATUS_SUCCESS)
					ret = lret;
			} else
				ret = VariationsAddVariation( doc3d, variations, variation );
		} else {
			for(val->value = val->start; val->value <= val->end; val->value += val->step) {
				if (index < (variation->variables.size() - 1)) {
					INT lret = DoRange(doc3d, variations, variation, index + 1);
					if (lret != LIBSTATUS_SUCCESS) {
						ret = lret;
						break;
					}
				} else {
					ret = VariationsAddVariation( doc3d, variations, variation );
					if (ret != LIBSTATUS_SUCCESS)
						break;
				}
			}
		}
	}

	return ret;
}

UINT CBatchExportApp::ComposeString( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR templ, VARIATION * variation, FIELD field )
{
	// variables format and conversion stuff
	enum CONVTYPE { CT_NONE = 0, CT_DOUBLE, CT_INT64, CT_UINT64 };
	typedef const struct {
		WCHAR sel;
		CONVTYPE type;
		LPCWSTR convspec;
	} CONVT;
	const CONVT CONV[] = {
		{ L'g', CT_DOUBLE, L"g" },		// first element is the default one
		{ L'd', CT_INT64,  L"I64d" },
		{ L'f', CT_DOUBLE, L"f" },
		{ L'G', CT_DOUBLE, L"G" },
		{ L'F', CT_DOUBLE, L"F" },
		{ L'e', CT_DOUBLE, L"e" },
		{ L'E', CT_DOUBLE, L"E" },
		{ L'i', CT_INT64,  L"I64i" },
		{ L'u', CT_UINT64, L"I64u" },
		{ L'x', CT_UINT64,  L"I64x" },
		{ L'X', CT_UINT64,  L"I64X" },
		/* must be the last */
		{ 0, CT_NONE, NULL }
	};

	UINT errkind = LIBSTATUS_ERR_INVARG;
	UINT errkindnv = LIBSTATUS_SUCCESS;
	BOOL delayed = FALSE;
	INT maxlen = 511;
	UINT ret = LIBSTATUS_SUCCESS;
	CString *strptr = NULL;
	switch(field) {
	case VP_DIR:
		strptr = &variation->dir;
		errkind = LIBSTATUS_BAD_DIR;
		errkindnv = LIBSTATUS_DIR_NO_VAR;
		delayed = (variation->options & VAROPT_DELAYED_COMPOSE_DIR) == VAROPT_DELAYED_COMPOSE_DIR;
		maxlen = 32766;
		break;
	case VP_NAME:
		strptr = &variation->name;
		errkind = LIBSTATUS_BAD_NAME;
		errkindnv = LIBSTATUS_NAME_NO_VAR;
		delayed = (variation->options & VAROPT_DELAYED_COMPOSE_NAME) == VAROPT_DELAYED_COMPOSE_NAME;
		break;
	case VP_COMMENTADD:
		strptr = &variation->commentadd;
		errkind = LIBSTATUS_BAD_COMMENTADD;
		errkindnv = LIBSTATUS_COMMENTADD_NO_VAR;
		delayed = (variation->options & VAROPT_DELAYED_COMPOSE_COMMENTADD) == VAROPT_DELAYED_COMPOSE_COMMENTADD;
		break;
	}
	if (strptr == NULL)
		return errkind;

	LPCWSTR templptr = templ;
	BOOL bvar = FALSE;

	strptr->Empty();
	while(*templptr) {
		size_t len = wcscspn(templptr, L"$");
		LPCWSTR next = templptr + len + 1;
		if (bvar) {
			if (len < 1)														// $ char
				strptr->AppendChar('$');
			else {
				if (templptr[0] == '@') {										// extended options
					LPWSTR endptr;
					ULONG index;
					LONG opt = 0;
					WCHAR kind = '\0';

					templptr++;
					len--;
					kind = toupper(templptr[0]);
					templptr++;
					len--;
					endptr = (LPWSTR)templptr + len; // init endptr to the end of the current pattern
					switch(kind) {
					case 'E': { // embodiment marking
							LPWSTR leptr;

							index = wcstoul(templptr, &leptr, 0);

							if (leptr == templptr)
								index = variation->embodiment;
							else {
								len -= leptr - templptr;
								templptr = leptr;
							}
							opt = ksVMFullMarking;
							if (len > 0) {
								if (*templptr != '.')
									return errkind;

								templptr++;
								len--;

								if (!ParseMarking(templptr, len, opt))
									return errkind;
							}
							if (!Document::EmbodimentMarkingGet( doc3d, index, opt, strptr, TRUE ))
								return errkindnv;
						}
						break;
					case 'N': { // embodiment name
							LPWSTR mname;
							LPWSTR leptr;

							index = wcstoul(templptr, &leptr, 0);
							if (templptr == leptr)
								index = variation->embodiment;
							if (leptr != endptr)
								return errkind;
							if (!Document::ModelEmbodimentNameGet(doc3d, index, mname))
								return errkindnv;
							strptr->Append(mname);
						}
						break;
					case 'V': { // variable value
							if (templptr == endptr)
								return errkind;

							BOOL comma = FALSE;
							while(len > 0) {			// skip spaces and comma (but only one) if any
								switch (*templptr) {
								case L',':
									if (comma)
										return errkind;
									comma = TRUE;
								case L' ':
								templptr++; len--;
								continue;
								}
								break;
							}

							DWORD bufsz = 0;
							CString varname(templptr, (INT)len);
							bufsz = GetEnvironmentVariable(varname, NULL, 0);
							if (bufsz != 0) {
								INT strlen = strptr->GetLength();
								LPWSTR bufptr = strptr->GetBufferSetLength(strlen + bufsz);
								if (bufptr == NULL)
									return LIBSTATUS_SYSERR_NO_MEM;
								GetEnvironmentVariable(varname, bufptr + strlen, bufsz);
								strptr->ReleaseBuffer();
							}
						}
						break;
					case 'D':
					case 'T': { // date/time
							INT const bufsz = 128;
							INT strlen = strptr->GetLength();
							LPWSTR bufptr = strptr->GetBufferSetLength(strlen + bufsz);
							if (bufptr == NULL)
								return LIBSTATUS_SYSERR_NO_MEM;
							CString dateformat(templptr, (INT)len);
							size_t ret = _wcsftime_l(bufptr + strlen, bufsz, dateformat, &variation->time, m_timelocale);
							strptr->ReleaseBuffer();
							if (ret == 0)
								return errkind;
						}
						break;
					case 'L': {
							LPCWSTR fname = doc3d->GetName();
							LPCWSTR fext = PathFindExtension(fname);
							if (fext)
								strptr->Append(fname, fext - fname);
							else
								strptr->Append(fname);
						}
						break;
// ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ add new features above ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ //
					default:
						templptr--; // defaulting to feature name, revert pointer
						len++;
					case 'F': { // feature name
							LPWSTR leptr;

							if (templptr == endptr)
								return errkind;

							index = wcstoul(templptr, &leptr, 0);
							if (leptr != endptr)
								return errkind;
							if (!Document::NFeatureNameGet( doc3d, index, strptr, TRUE, variation->embodiment ))
								return errkindnv;
							break;
						}
						break;
					}
				} else {						// document variable
					CONVT *conv = CONV;
					if (*templptr == '^') { templptr++; len--; } // FIXME backward compatibility
					if (*templptr == L'%') {	// alternative specifier
						templptr++; len--;
						conv = NULL;
					}
					LPCWSTR modptr = templptr;
					size_t modlen = wcsspn(modptr, L" #+-.0123456789");
					templptr += modlen; len -= modlen;
					if (len == 0)
						return errkind;
					if (conv == NULL) {			// alternative specifier
						WCHAR convsel = *templptr;
						templptr++; len--;
						for(conv = CONV; conv->sel && (conv->sel != convsel); conv++);
						if (!conv->sel)			// illegal specifier
							return errkind;
					}
					BOOL comma = FALSE;
					while(len > 0) {
						switch (*templptr) {
						case L',':
							if (comma)
								return errkind;
							comma = TRUE;
						case L' ':
							templptr++; len--;
							continue;
						}
						break;
					}
					if (!Document::IsVariableNameValid(templptr, (LONG)len)) // variable name check
						return errkind;
					size_t cslen = wcslen(conv->convspec);
					size_t fsz = modlen + cslen + len + 5; // + % + modifier + convspec + \0\0 + varname + \0\0
					LPWSTR format = new WCHAR[fsz];
					if (format == NULL)
						return LIBSTATUS_SYSERR_NO_MEM;
					wcscpy_s(format, fsz, L"%");
					if (modlen > 0)
						wcsncpy_s(format + 1, fsz - 1, modptr, modlen); // + %
					LPWSTR cs = format + modlen + 1; // + % + modifier
					LPWSTR varname = format + fsz - len - 2; // - varname - \0\0
					wcsncpy_s(varname, len + 2, templptr, len);
					size_t i = 0, varsiz = 0;
					if (!delayed)
						varsiz = variation->variables.size();
					for(; i < varsiz; i++) {
						if (wcscmp(variation->variables[i].name, varname) == 0) {
							auto *valptr = &variation->variables[i].values[variation->variables[i].curval];
							if (!valptr->expr.IsEmpty()) {
								wcscpy_s(cs, cslen + 2, L"s");
								strptr->AppendFormat(format, valptr->expr);
							} else {
								wcscpy_s(cs, cslen + 2, conv->convspec);
								switch(conv->type) {
								case CT_DOUBLE: strptr->AppendFormat(format, valptr->value); break;
								case CT_INT64: strptr->AppendFormat(format, (INT64)valptr->value); break;
								case CT_UINT64: strptr->AppendFormat(format, (UINT64)valptr->value); break;
								}
							}
							break;
						}
					}
					if (i == varsiz) { // variable not found in variables or is a delayed composition, try to get it from the document
						DOUBLE value = 0.0;
						if (Document::VariableValueGet(doc3d, varname, value, variation->embodiment)) {
							wcscpy_s(cs, cslen + 2, conv->convspec);
							switch(conv->type) {
							case CT_DOUBLE: strptr->AppendFormat(format, value); break;
							case CT_INT64: strptr->AppendFormat(format, (INT64)value); break;
							case CT_UINT64: strptr->AppendFormat(format, (UINT64)value); break;
							}
						} else
							ret |= errkindnv;
					}
					delete[] format;
				}
			}
		} else if (len > 0)						// plain text
			strptr->Append(templptr, (INT)len);
		if (templptr[len] == '\0')
			break;
		bvar = !bvar;
		templptr = next;
	}
	strptr->Trim(' ');
	if (strptr->GetLength() > maxlen)
		strptr->Truncate(maxlen);
	if (bvar)
		ret |= errkind;
	return ret;
}

UINT CBatchExportApp::ComposeDir( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR dirtempl, VARIATION * variation )
{
	INT ret = LIBSTATUS_SUCCESS;
	if ((ret = ComposeString(doc3d, dirtempl, variation, VP_DIR)) != LIBSTATUS_SUCCESS)
		return ret;
	variation->dir.Trim(' ');
	LPWSTR real = NULL;
	LPWSTR docpath = (LPWSTR)doc3d->GetPath();
	if (wcslen(docpath) == 0) {
		docpath = new WCHAR[MAX_PATH];
		if (docpath == NULL)
			return LIBSTATUS_SYSERR_NO_MEM;
		HRESULT hres = SHGetFolderPath(NULL, CSIDL_PERSONAL, NULL, SHGFP_TYPE_CURRENT, docpath);
		if (hres != S_OK) {
			delete[] docpath;
			return LIBSTATUS_DIR_UNKNOWN;
		}
		real = Utils::AbsPath((LPCWSTR)variation->dir, docpath, FALSE, '_');
		delete[] docpath;
	} else
		real = Utils::AbsPath((LPCWSTR)variation->dir, docpath, FALSE, '_');
	if (real == NULL)
		return LIBSTATUS_BAD_DIR;
	variation->dir.SetString(real);
	delete[] real;

	/*if (variation->dir.GetLength() >= MAX_PATH)
		return LIBSTATUS_DIR_LENGTH;*/

	return ret;
}

#define VAR_STEP_DEFAULT 1.0
#define ABS(a) ((a)<0?-(a):(a))
UINT CBatchExportApp::ParseVariables( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varstr, VARIATION * variation, LPCWSTR * endptr )
{
	VARIABLES *variables = &variation->variables;
	const VARVALUE val = { L"", 0.0, 0.0, 0.0, VAR_STEP_DEFAULT };
	VARVALUE *valptr = NULL;
	const VARIABLE var = { L"", VOPT_NONE, 0 };
	VARIABLE *varptr = NULL;
	LPCWSTR strptr = varstr;
	enum {
		PVS_INIT = 0,
		PVS_NAME,
		PVS_GOTNAME,
		PVS_OPTS,
		PVS_GOTOPTS,
		PVS_VAR,
		PVS_COL1,
		PVS_END,
		PVS_COL2,
		PVS_STEP,
		PVS_COMR,
		PVS_EXPR,
		PVS_INQ,
		PVS_COME
	};
	UINT state = PVS_INIT;

	variables->clear();
	do {
		if (endptr)
			*endptr = strptr;
		WCHAR ch = *strptr;
		if (state == PVS_EXPR) {
			switch (ch) {
			case '\0':
				return LIBSTATUS_BAD_VARS;
			case '\"':
				break;
			// TODO: expression char checks?
			default:
				valptr->expr += ch;
				continue;
			}
		}
		switch(ch) {
		case ' ':
		case '\t':
			switch(state) {
			case PVS_NAME:
				state = PVS_GOTNAME;
				break;
			case PVS_OPTS:
				state = PVS_GOTOPTS;
				break;
			}
			continue;
		case '\"':
			switch(state) {
			case PVS_VAR:
				state = PVS_EXPR;
				continue;
			case PVS_EXPR:
				state = PVS_COME;
				continue;
			}
			return LIBSTATUS_BAD_VARS;
		case '|': // option: local
			if ((state < PVS_NAME) || (state > PVS_OPTS))
				return LIBSTATUS_BAD_VARS;
			varptr->opt |= VOPT_LOCAL;
			state = PVS_OPTS;
			break;
		case '=': // name separator
			if ((state < PVS_NAME) || (state > PVS_GOTOPTS))
				return LIBSTATUS_BAD_VARS;
			if (Document::EmbodimentVariable(doc3d, varptr->name, varptr->opt & VOPT_LOCAL?variation->embodiment:Document::EMBODIMENT_TOP) == NULL)
				return LIBSTATUS_VARS_NO_VAR;
			varptr->values.push_back(val); // add first value
			valptr = &varptr->values.back();
			if (valptr == NULL)
				return LIBSTATUS_SYSERR_NO_MEM;
			state = PVS_VAR;
			break;
		case ':':
			switch(state) {
			case PVS_COL1:
			case PVS_COL2:
				state++;
				continue;
			}
			return LIBSTATUS_BAD_VARS;
		case ',': // values separator
			switch(state) {
			case PVS_COMR:
			case PVS_COME:
			case PVS_COL1:
			case PVS_COL2:
				break;
			default:
				return LIBSTATUS_BAD_VARS;
			}
		case '\0':
		case ';': // variables separators
			switch(state) {
			case PVS_INIT:
				continue;
			case PVS_COL1:
				valptr->end = valptr->start;
			case PVS_COL2:
			case PVS_COMR:
				if (ABS((valptr->end - valptr->start)/valptr->step) > 1000.0)
					return LIBSTATUS_BAD_VARS;
				if ((valptr->end < valptr->start) && (valptr->step > 0))
					valptr->step = -valptr->step;
				valptr->value = valptr->start;
				break;
			case PVS_COME:
				valptr->expr.Trim();
				if (!valptr->expr.IsEmpty())
					break;
			default:
				return LIBSTATUS_BAD_VARS;
			}
			valptr = NULL;
			if (ch == ',') { // try next value
				varptr->values.push_back(val); // add new value
				valptr = &varptr->values.back();
				if (valptr == NULL)
					return LIBSTATUS_SYSERR_NO_MEM;
				state = PVS_VAR;
				break;
			}
			state = PVS_INIT;
			varptr = NULL;
			break;
		default:
			switch (state) {
			case PVS_INIT:
			case PVS_NAME:
				if (varptr == NULL) { // add new variable
					if (!Document::IsVariableNameCharValid(ch, TRUE)) // first char
						return LIBSTATUS_BAD_VARS;
					variables->push_back(var);
					varptr = &variables->back();
					if (varptr == NULL)
						return LIBSTATUS_SYSERR_NO_MEM;
				} else if (!Document::IsVariableNameCharValid(ch)) // following chars
					return LIBSTATUS_BAD_VARS;
				varptr->name += ch;
				state = PVS_NAME;
				break;
			case PVS_VAR:
				if (!Utils::ParseDouble(strptr, strptr, valptr->start, m_numlocale))
					return LIBSTATUS_BAD_VARS;
				strptr--;  // rewind one consumed char
				state = PVS_COL1;
				break;
			case PVS_END:
				if (!Utils::ParseDouble(strptr, strptr, valptr->end, m_numlocale))
					return LIBSTATUS_BAD_VARS;
				strptr--;
				state = PVS_COL2;
				break;
			case PVS_STEP:
				if (!Utils::ParseDouble(strptr, strptr, valptr->step, m_numlocale))
					return LIBSTATUS_BAD_VARS;
				strptr--;
				state = PVS_COMR;
				break;
			default:
				return LIBSTATUS_BAD_VARS;
			}
		}
	} while (*strptr++);
	if (state != PVS_INIT)
		return LIBSTATUS_BAD_VARS;

	return LIBSTATUS_SUCCESS;
}

UINT CBatchExportApp::ComposeVariables( ksAPI7::IKompasDocument3DPtr & doc3d, VARIABLES * variables, CString * varstr, BOOL values )
{
	UINT ret = LIBSTATUS_SUCCESS;
	varstr->Empty();
	for(size_t vari = 0; vari < variables->size(); vari++) {  // somehow it's faster than using iterators
		auto *varptr = &variables->at(vari);
		if (varptr->values.size() == 0)
			return LIBSTATUS_BAD_VARS;
		if (vari > 0)
			varstr->AppendChar(';');
		varstr->Append(varptr->name);
		if ((varptr->opt & VOPT_LOCAL) == VOPT_LOCAL)
			varstr->AppendChar('|');
		varstr->AppendChar('=');
		for(size_t vali = 0; vali < varptr->values.size(); vali++) {
			auto *valptr = &varptr->values.at(vali);
			if (vali > 0)
				varstr->AppendChar(',');
			if (!valptr->expr.IsEmpty()) {
				varstr->AppendChar('"');
				varstr->Append(valptr->expr);
				varstr->AppendChar('"');
			} else if (values) {
				varstr->AppendFormat(L"%g", valptr->value);
			} else {
				varstr->AppendFormat(L"%g", valptr->start);
				if (valptr->end != valptr->start)
					varstr->AppendFormat(L":%g", valptr->end);
				if (valptr->step != VAR_STEP_DEFAULT)
					varstr->AppendFormat(L":%g", valptr->step);
			}
		}
	}

	return ret;
}

UINT CBatchExportApp::ParseOptions( LPCWSTR optstr, ULONG * options )
{
	*options = VAROPT_DEFAULT;

	size_t optlen = wcslen(optstr);
	if (optlen < 1)
		return LIBSTATUS_SUCCESS;
	LPWSTR optdup = new WCHAR[optlen + 1];
	if (optdup == NULL)
		return LIBSTATUS_SYSERR_NO_MEM;
	wcscpy_s(optdup, optlen + 1, optstr);
	_wcslwr_s(optdup, optlen + 1);
	LPWSTR optptr = optdup;
	LPWSTR next = NULL;

	BOOL last = FALSE;
	UINT ret = LIBSTATUS_SUCCESS;
	while(*optptr) {
		size_t len = wcscspn(optptr, L",:; \t");
		next = optptr + len + 1;
		if (optptr[len] == '\0')
			last = TRUE;
		else
			optptr[len] = '\0';
		if (len > 0) {
			const OPTIONMAP *opt;
			for(opt = optionmap; opt->name && (Utils::Matches(optptr, opt->name) != 0); opt++);
			if (opt->name != NULL) {
				*options &= ~opt->mask;
				if (opt->val)
					*options |= opt->val;
			} else
				ret |= LIBSTATUS_BAD_OPTS;
		}
		if (last)
			break;
		optptr = next;
	}

	delete[] optdup;
	return ret;
}

void CBatchExportApp::ComposeOptions( ULONG options, CString * optstr, BOOL omitdefault )
{
	const OPTIONMAP *opt;

	optstr->Empty();
	for(opt = optionmap; opt->name; opt++) {
		if ((omitdefault && (opt->val == (VAROPT_DEFAULT & opt->mask))) ||
			(opt->val != (options & opt->mask)))
			continue;
		if (!optstr->IsEmpty())
			optstr->Append(L" ");
		optstr->Append(opt->name);
	}
}

BOOL CBatchExportApp::ParseMarking( LPCWSTR str, size_t slen, LONG & val )
{
    if ((str == NULL) || (slen == 0))
        return TRUE;

    typedef const struct {
        LPCWSTR name;
        UINT len;
        INT val;
    } markopts_t;
#define ADD_MOPT(opt) { L#opt, _countof(#opt) - 1/*\0*/, opt }
	// option with the highest value must be the last
    markopts_t markopts[] = { ADD_MOPT(ksVMFullMarking), ADD_MOPT(ksVMBaseMarking), ADD_MOPT(ksVMEmbodimentNumber),
                              ADD_MOPT(ksVMAdditionalNumber), ADD_MOPT(ksVMCode), { NULL, 0, 0 } };

    LONG lval = 0;
    LPCWSTR kwptr = NULL;
    size_t kwlen = 0;
    INT opOR = -1;
    do {
        WCHAR ch = *str;
        switch(ch) {
        case '\0':
			return FALSE;
        case '|':
            if (opOR != FALSE)
                return FALSE;
			opOR = TRUE;
		case ' ':
        case '\t':
            break;
        default:
            if (kwptr == NULL) {
				if (opOR == FALSE)
					return FALSE;
                kwptr = str;
				opOR = FALSE;
			}
			if (slen > 1)
				continue;
			str++;
        }
		if (kwptr != NULL) {
			kwlen = str - kwptr;

			markopts_t *i;
			for(i = markopts; i->name != NULL; i++) {
				if (kwlen != i->len)
					continue;
				if (_wcsnicmp(i->name, kwptr, i->len) == 0) {
					lval |= i->val;
					break;
				}
			}
			if (i->name == NULL) {
				LPWSTR leptr;
				lval |= wcstol(kwptr, &leptr, 0);
				if ((leptr - kwptr) != kwlen)
					return FALSE;
			}

			kwptr = NULL;
			kwlen = 0;
		}
    } while(--slen && ++str);

    if (opOR == TRUE)
        return FALSE;

	if (lval > ((markopts[_countof(markopts)-2].val << 1) - 1))
		return FALSE;
	if (lval < ksVMFullMarking)
		val = ksVMFullMarking;
	else
		val = lval;
	return TRUE;
}

BOOL CBatchExportApp::GetAttributeLib( LPWSTR & lib )
{
	LPWSTR mpath = NULL;
	ksAttributeTypeT attrType;
	if (!ksGetAttrTypeW (VARTYPENUM, mpath, &attrType)) {
		ResultNULL();
		const SIZE_T bufsz = 32768;
		mpath = new WCHAR[bufsz];
		if (mpath == NULL)
			return FALSE;
		SIZE_T nameidx = 0;
		if (!Utils::GetModulePathName(mpath, bufsz - _countof(VARTYPELIB) + 1, &nameidx)) {
			MessageBox((HWND)GetHWindow(), L"Не удалось получить путь к библиотеке атрибутов " VARTYPELIB L".", theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONEXCLAMATION);
			delete[] mpath;
			return FALSE;
		}

		wcscpy_s(&mpath[nameidx], bufsz - nameidx, VARTYPELIB);
		if (!ksGetAttrTypeW (VARTYPENUM, mpath, &attrType)) {
			ResultNULL();
			delete[] mpath;
			CString message;
			message.Format(L"В документе отсутствует тип атрибута с UID " UID L", библиотека атрибутов %s не найдена или не содержит тип атрибута с UID" UID L".", mpath);
			MessageBox((HWND)GetHWindow(), message, theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONEXCLAMATION);
			return FALSE;
		}
		CString message;
		message.Format(L"В документе отсутствует тип атрибута с UID " UID L", использовать тип из библиотеки %s?", mpath);
		if (MessageBox((HWND)GetHWindow(), message, theApp.LoadStr(IDR_LIBID), MB_YESNO|MB_ICONQUESTION) != IDYES) {
			delete[] mpath;
			return FALSE;
		}
	}

	lib = mpath;
	return TRUE;
}

BOOL CBatchExportApp::StrStatus( CString & str, UINT status, BOOL append )
{
	UINT field = status;
	if (field == LIBSTATUS_SUCCESS)
		return FALSE;
	if (Utils::ComStrStatus(str, field, append)) // check for common errors
		return TRUE;

	// it is not a common error, so field contains our own information only
	if (field & LIBSTATUS_EXCLUSIVE) { // exclusive status
		LPCWSTR strptr = NULL;
		switch(field) {
		// informational
		case LIBSTATUS_INFO_DONE:				strptr = L"Готово";											break;
		case LIBSTATUS_INFO_TESTOK:				strptr = L"Тест пройден";									break;
		// warnings
		case LIBSTATUS_WARN_FORMAT_UNSUP:		strptr = L"Выбранный формат не поддерживается";				break;
		// runtime errors
		case LIBSTATUS_ERR_SET_EMBODIMENT:		strptr = L"Ошибка выбора исполнения";						break;
		case LIBSTATUS_ERR_SET_NAME:			strptr = L"Ошибка установки имени";							break;
		case LIBSTATUS_ERR_SET_COMMENT:			strptr = L"Ошибка добавления комментария";					break;
		case LIBSTATUS_ERR_SET_VARIABLE:		strptr = L"Ошибка установки переменной";					break;
		case LIBSTATUS_ERR_DOCUMENT_REBUILD:	strptr = L"Ошибка перестроения документа";					break;
		case LIBSTATUS_ERR_DOCUMENT_INVALID:	strptr = L"Документ стал недействительным";					break;
		case LIBSTATUS_ERR_DISCARDED:			strptr = L"Вариация отвергнута";							break;
		case LIBSTATUS_ERR_VARIATIONS_EXCLUDED:	strptr = L"Все вариации были исключены";					break;
		case LIBSTATUS_ERR_NO_ATTRS:			strptr = L"В документе не найдено атрибутов VARIATIONS с UID "UID; break;
		case LIBSTATUS_ERR_NO_EMBODIMENTS:		strptr = L"Документ не имеет исполнений";					break;
		case LIBSTATUS_ERR_NO_VARIATIONS:		strptr = L"Не создано ни одной вариации";					break;
		case LIBSTATUS_ERR_CURRENT_TIME:		strptr = L"Не удалось получить системное время или отказ localtime()"; break;
		}
		if (strptr == NULL)
			return FALSE; // string doesn't changed
		if (append)
			str.Append(strptr);
		else
			str.SetString(strptr);
		return TRUE;
	}
	// non exclusive input errors
	if (!append) str.Empty();
	BOOL comma = FALSE;
	while(field) {
		// invaliddata @ https://stackoverflow.com/questions/3142867/finding-bit-positions-in-an-unsigned-32-bit-integer
		UINT temp = field & -(INT)field;  //extract least significant bit on a 2s complement machine, MSb already tested
		field ^= temp;  // toggle the bit off
		//now you could have a switch statement or bunch of conditionals to test temp
		//or get the index of the bit and index into a jump table, etc.
		if (comma)
			str += L", ";
		switch(temp) {
		case LIBSTATUS_BAD_EMBODIMENT:	str += L"неправильное определение исполнения";				break;
		case LIBSTATUS_BAD_DIR:			str += L"неправильное определение директории";				break;
		case LIBSTATUS_DIR_UNKNOWN:		str += L"невозможно получить целевую директорию";			break;
		case LIBSTATUS_DIR_NO_VAR:
		case LIBSTATUS_NAME_NO_VAR:
		case LIBSTATUS_COMMENTADD_NO_VAR:
		case LIBSTATUS_VARS_NO_VAR:		str += L"Переменная или элемент не найдены в исполнении";	break;
		case LIBSTATUS_BAD_NAME:		str += L"неправильное определение имени";					break;
		case LIBSTATUS_BAD_COMMENTADD:	str += L"неправильное определение коментария";				break;
		case LIBSTATUS_BAD_VARS:		str += L"неправильное определение переменных";				break;
		case LIBSTATUS_BAD_OPTS:		str += L"неправильное определение опций";					break;
		default: str.TrimRight(L", "); /* normally shouldn't happen */								break;
		}
		comma = TRUE;
	}

	return TRUE;
}

CString CBatchExportApp::LoadStr( UINT strID )
{
	CString str;
	str.LoadStringW(m_hInstance, strID);
	return str;
}

SIZE_T CBatchExportApp::LoadStr( UINT strID, LPWSTR buf, SIZE_T bufsz )
{
	return LoadString(m_hInstance, strID, buf, bufsz);
}

void CBatchExportApp::EnableTaskAccess( BOOL enable )
{
	m_kompas->ksEnableTaskAccess(enable);
}

// private
BOOL CBatchExportApp::GetKompas() 
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
	StrStatus(err, LIBSTATUS_SYSERR|dwErr, TRUE);

	MessageBox((HWND)GetHWindow(), err, theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONERROR);

	return FALSE;
}

// static
static void EmbodimentIndexesShow( Document::EMBODIMENTNAMES & names )
{
	size_t count = names.size();
	CString msg;
	if (count > 0) {
		size_t index = 0;
		while (index < count) {
			msg.Empty();
			for (; index < count; index++) {
				msg.AppendFormat(L"% 5d: %s\t%s\n", index, names[index].name, names[index].number);
				if ((index % 20) == 19) {
					msg.Append(L"\nДалее?");
					INT ret = MessageBox((HWND)GetHWindow(), (LPCWSTR)msg, L"Перечисление исполнений", MB_YESNO|MB_DEFBUTTON1|MB_ICONINFORMATION);
					msg.Empty();
					if (ret != IDYES)
						return;
				}
			}
			if (!msg.IsEmpty())
				MessageBox((HWND)GetHWindow(), (LPCWSTR)msg, L"Перечисление исполнений", MB_OK|MB_ICONINFORMATION);
		}
	}
}

static void FeatureIndexesShow( Document::FEATURENAMES & names )
{
	size_t count = names.size();
	CString msg;
	if (count > 0) {
		size_t index = 0;
		while (index < count) {
			msg.Empty();
			for (; index < count; index++) {
				msg.AppendFormat(L"% 5d: %s\n", index, names[index].name);
				if ((index % 20) == 19) {
					msg.Append(L"\nДалее?");
					INT ret = MessageBox((HWND)GetHWindow(), (LPCWSTR)msg, L"Перечисление элементов", MB_YESNO|MB_DEFBUTTON1|MB_ICONINFORMATION);
					msg.Empty();
					if (ret != IDYES)
						return;
				}
			}
			if (!msg.IsEmpty())
				MessageBox((HWND)GetHWindow(), (LPCWSTR)msg, L"Перечисление элементов", MB_OK|MB_ICONINFORMATION);
		}
	}
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
			case  IDR_BATCHEXPORT: theApp.BatchExport( doc3d ); break;
			case  IDR_EMBENUM:
				{
					Document::EMBODIMENTNAMES list;
					if (Document::EmbodimentNameEnum( doc3d, list ))
						EmbodimentIndexesShow( list );
					else
						MessageBox((HWND)GetHWindow(), L"Ошибка перечисления исполнений", L"Перечисление исполнений", MB_OK|MB_ICONERROR);
				}
				break;
			case  IDR_FEAENUM:
				{
					Document::FEATURENAMES list;
					if (Document::FeatureNameEnum( doc3d, list, Document::EMBODIMENT_CURRENT ))
						FeatureIndexesShow( list );
					else
						MessageBox((HWND)GetHWindow(), L"Ошибка перечисления элементов", L"Перечисление элементов", MB_OK|MB_ICONERROR);
				}
				break;
			case IDR_VAREXP:
				{
					UINT res = theApp.VariationsExport( doc3d, TRUE );
					if (res != LIBSTATUS_SUCCESS && res != LIBSTATUS_ERR_ABORTED) {
						CString errstr(L"Ошибка экспорта шаблонов: ");
						CBatchExportApp::StrStatus(errstr, res, TRUE);
						MessageBox((HWND)GetHWindow(), errstr, L"Экспорт шаблонов", MB_OK|MB_ICONERROR);
					}
				}
				break;
			case IDR_VARIMP:
				{
					UINT res = theApp.VariationsImport( doc3d, TRUE );
					if (res != LIBSTATUS_SUCCESS && res != LIBSTATUS_ERR_ABORTED) {
						CString errstr(L"Ошибка импорта шаблонов: ");
						CBatchExportApp::StrStatus(errstr, res, TRUE);
						MessageBox((HWND)GetHWindow(), errstr, L"Импорт шаблонов", MB_OK|MB_ICONERROR);
					}
				}
		}
		theApp.EnableTaskAccess(TRUE);
	} else if (theApp.hasAPI())
		MessageBox((HWND)GetHWindow(), theApp.LoadStr(IDS_NODOC), theApp.LoadStr(IDR_LIBID), MB_OK|MB_ICONEXCLAMATION);
}
