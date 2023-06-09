////////////////////////////////////////////////////////////////////////////////
//
// BatchExport.h
// BatchExport library
// Alex Fedorov, 2018
//
////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_BATCHEXPORT_H__8FC6CC92_1590_41F3_83BF_4EC6CA01F2C8__INCLUDED_)
#define AFX_BATCHEXPORT_H__8FC6CC92_1590_41F3_83BF_4EC6CA01F2C8__INCLUDED_

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "..\common\common.h"   // common stuff
#include "..\common\document.h" // document stuff
#include "resource.h"           // main symbols

#define VARTYPENUM 289001038138.0
#define VARTYPELIB L"BatchExport.lat"

#define STRINGIZE2(s) L#s
#define STRINGIZE(s) STRINGIZE2(s)
#define UID STRINGIZE(VARTYPENUM)

// project specific errors
typedef enum {
	LIBSTATUS_BAD_EMBODIMENT		= 0x1,		  // non-exclusive input errors
	LIBSTATUS_EMBODIMENT_MASK		= LIBSTATUS_BAD_EMBODIMENT,
	LIBSTATUS_BAD_DIR				= 0x2,
	LIBSTATUS_DIR_UNKNOWN			= 0x4,
	LIBSTATUS_DIR_NO_VAR			= 0x8,
	LIBSTATUS_DIR_MASK				= LIBSTATUS_BAD_DIR | LIBSTATUS_DIR_UNKNOWN | LIBSTATUS_DIR_NO_VAR,
	LIBSTATUS_BAD_NAME				= 0x10,
	LIBSTATUS_NAME_NO_VAR			= 0x20,
	LIBSTATUS_NAME_MASK				= LIBSTATUS_BAD_NAME | LIBSTATUS_NAME_NO_VAR,
	LIBSTATUS_BAD_COMMENTADD		= 0x40,
	LIBSTATUS_COMMENTADD_NO_VAR		= 0x80,
	LIBSTATUS_COMMENTADD_MASK		= LIBSTATUS_BAD_COMMENTADD | LIBSTATUS_COMMENTADD_NO_VAR,
	LIBSTATUS_BAD_VARS				= 0x100,
	LIBSTATUS_VARS_NO_VAR			= 0x200,
	LIBSTATUS_VARS_MASK				= LIBSTATUS_BAD_VARS | LIBSTATUS_VARS_NO_VAR,
	LIBSTATUS_BAD_OPTS				= 0x400,
	LIBSTATUS_OPTS_MASK				= LIBSTATUS_BAD_OPTS,
	// NOT including MASK values
	LIBSTATUS_USER_ERR_MAX			= LIBSTATUS_BAD_OPTS,

	LIBSTATUS_ERR_SET_EMBODIMENT		= LIBSTATUS_ERR | 0x1,
	LIBSTATUS_ERR_SET_NAME				= LIBSTATUS_ERR | 0x2,
	LIBSTATUS_ERR_SET_COMMENT			= LIBSTATUS_ERR | 0x3,
	LIBSTATUS_ERR_SET_VARIABLE			= LIBSTATUS_ERR | 0x4,
	LIBSTATUS_ERR_DOCUMENT_REBUILD		= LIBSTATUS_ERR | 0x5,
	LIBSTATUS_ERR_DOCUMENT_INVALID		= LIBSTATUS_ERR | 0x6,
	LIBSTATUS_ERR_DISCARDED				= LIBSTATUS_ERR | 0x7,
	LIBSTATUS_ERR_VARIATIONS_EXCLUDED	= LIBSTATUS_ERR | 0x8,
	LIBSTATUS_ERR_NO_ATTRS				= LIBSTATUS_ERR | 0x9,
	LIBSTATUS_ERR_NO_EMBODIMENTS		= LIBSTATUS_ERR | 0xA,
	LIBSTATUS_ERR_NO_VARIATIONS			= LIBSTATUS_ERR | 0xB,
	LIBSTATUS_ERR_CURRENT_TIME			= LIBSTATUS_ERR | 0xC,

	LIBSTATUS_WARN_FORMAT_UNSUP			= LIBSTATUS_WARN | 0x1,

	LIBSTATUS_INFO_DONE					= LIBSTATUS_INFO | 0x1,
	LIBSTATUS_INFO_TESTOK				= LIBSTATUS_INFO | 0x2
};

class CBatchExportApp : public CWinApp
{
public:
	// enums
	// variations data indexes
	// on chage update VariationGet, VariationsSave, GridViewDlg::OnInitDialog
	typedef enum {
		VP_EMBODIMENT = 0,
		VP_DIR = 1,
		VP_NAME = 2,
		VP_COMMENTADD = 3,
		VP_VARIABLES = 4,
		VP_OPTIONS = 5,
		VP_STATUS = 6,
		FIELD_COUNT = 7
	} FIELD;

	typedef enum {
		// on including additional options, update optionmap[] in BatchExport.cpp accordingly
		// options that not included in optionmap[] are internal hidden options, inaccessible to user
		VAROPT_INCLUDE		 = 0x80000000,
		VAROPT_CURVES		 = 0x40000000,
		VAROPT_DELAYED_COMPOSE_DIR        = 0x100,
		VAROPT_DELAYED_COMPOSE_NAME       = 0x200,
		VAROPT_DELAYED_COMPOSE_COMMENTADD = 0x400,
		VAROPT_DELAYED_COMPOSE_MASK       = 0x700,
		// format is exclusive and must occupy LS bits
		VAROPT_FORMAT_MASK	 = 0xFF,
		VAROPT_FORMAT_NATIVE = 0x00,
		VAROPT_FORMAT_AP203	 = 0x01,
		VAROPT_FORMAT_AP214	 = 0x02,
		VAROPT_FORMAT_STL    = 0x03,
		VAROPT_FORMAT_VRML   = 0x04,
		// when adding support for additional formats, set VAROPT_FORMAT_MAX properly and update formatexts[] and optionmap[] in BatchExport.cpp
		VAROPT_FORMAT_MAX	 = VAROPT_FORMAT_VRML,
		VAROPT_DEFAULT		 = VAROPT_INCLUDE | VAROPT_CURVES | VAROPT_FORMAT_AP214
	} VAROPTION;

	typedef enum {
		VOPT_NONE			 = 0,
		VOPT_LOCAL			 = 1, // current embodiment variable
		VOPT_MAX			 = VOPT_LOCAL
	} VOPTION;

	typedef struct {
		CString expr;
		DOUBLE value;
		DOUBLE start;
		DOUBLE end;
		DOUBLE step;
	} VARVALUE;
	typedef std::vector <VARVALUE> VARVALUES;

	typedef struct {
		CString name;
		ULONG opt;
		ULONG curval;
		VARVALUES values;
	} VARIABLE;
	typedef std::vector <VARIABLE> VARIABLES;

	// typedefs
	typedef struct {
		reference AttrPtr;
		LONG AttrRow;
		LONG embodiment;
		CString dir;
		CString name;
		CString commentadd;
		CString varstr;
		CString optstr;
		VARIABLES variables;
		ULONG options;
		ULONG status;
		struct tm time;
	} VARIATION;
	typedef std::vector <VARIATION> VARIATIONS;

	// init
	CBatchExportApp();
	~CBatchExportApp();

	virtual BOOL InitInstance();
	virtual BOOL ExitInstance();

	// document manage
	ksAPI7::IKompasDocument3DPtr GetCurrentDocument();
	UINT PreparePath( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATION * variation, CString & patname );
	UINT SaveAs( ksAPI7::IKompasDocument3DPtr & doc3d, CString & pathname, const VARIATION * variation );
	UINT ExportFile( ksAPI7::IKompasDocument3DPtr & doc3d, CString & pathname, const VARIATION * variation );
	BOOL BatchExport( ksAPI7::IKompasDocument3DPtr & doc3d );

	BOOL VariableSet( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varname, VARIABLE * var, LONG embodiment = Document::EMBODIMENT_TOP );

	// variation
	LONG VariationsGetRowCount( reference pAttr );
	BOOL VariationGet( reference pAttr, LONG row, ksDynamicArrayPtr & valarray );
	BOOL VariationSet( reference pAttr, LONG row, ksDynamicArrayPtr & valarray );
	BOOL VariationUpdateTime( tm * ctime );
	BOOL VariationUpdateTime( VARIATION * variation );
	reference VariationsCreateAttr( ksAPI7::IKompasDocument3DPtr & doc3d, LPWSTR lib );
	UINT VariationsSave( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations );
	UINT VariationsGetTemplates( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations );
	UINT VariationsEnum( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations );
	UINT VariationsExport( ksAPI7::IKompasDocument3DPtr & doc3d, BOOL file = TRUE );
	UINT VariationsImport( ksAPI7::IKompasDocument3DPtr & doc3d, BOOL file = TRUE );
	UINT VariationsAddVariation( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations, VARIATION * variation );
	UINT DoRange( ksAPI7::IKompasDocument3DPtr & doc3d, VARIATIONS * variations, VARIATION * variation, ULONG index );

	// input handling
	UINT ComposeString( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR templ, VARIATION * variation, FIELD field );
	UINT ComposeDir( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR dirtempl, VARIATION * variation );
	UINT ParseVariables( ksAPI7::IKompasDocument3DPtr & doc3d, LPCWSTR varstr, VARIATION * variation, LPCWSTR * endptr = NULL );
	UINT ComposeVariables( ksAPI7::IKompasDocument3DPtr & doc3d, VARIABLES * variables, CString * varstr, BOOL values = FALSE );
	static UINT ParseOptions( LPCWSTR optstr, ULONG * options );
	static void ComposeOptions( ULONG options, CString * optstr, BOOL omitdefault = TRUE );

	// utility
	// static
	static BOOL ParseMarking( LPCWSTR str, size_t slen, LONG & val );
	static BOOL GetAttributeLib( LPWSTR & lib );
	static BOOL StrStatus( CString & str, UINT status, BOOL append = FALSE );

	BOOL hasAPI() { return (m_kompas != NULL) && (m_newKompasAPI != NULL); };
	CString LoadStr( UINT strID );
	SIZE_T LoadStr( UINT strID, LPWSTR buf, SIZE_T bufsz );
	void EnableTaskAccess( BOOL enable );

	// members
	static const LPCWSTR fieldnames[FIELD_COUNT];
	static const LPCWSTR formatexts[VAROPT_FORMAT_MAX];

private:
	// methods
	BOOL GetKompas();

	// members
	ksAPI7::IApplicationPtr m_newKompasAPI;               
	KompasObjectPtr m_kompas;
	_locale_t m_numlocale;  // explicit numeric locale to stricten radix character
	_locale_t m_timelocale; // explicit time locale for _wcsftime_l
	UINT m_overwritefiles;
	UINT m_testmode;
	UINT m_quick;
};

extern CBatchExportApp theApp;

#endif // !defined(AFX_BATCHEXPORT_H__8FC6CC92_1590_41F3_83BF_4EC6CA01F2C8__INCLUDED_)
