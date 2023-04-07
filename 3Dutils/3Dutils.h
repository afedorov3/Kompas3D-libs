////////////////////////////////////////////////////////////////////////////////
//
// 3Dutils.h - Библиотека на Visual C++
//
////////////////////////////////////////////////////////////////////////////////
#if !defined(AFX_3DUTILS_H__8FC6CC92_1590_41F3_83BF_4EC6CA01F2C8__INCLUDED_)
#define AFX_3DUTILS_H__8FC6CC92_1590_41F3_83BF_4EC6CA01F2C8__INCLUDED_

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "..\common\common.h"   // common stuff
#include "..\common\document.h" // document stuff
#include "resource.h"           // main symbols

#define DEFCOLOR RGB(144, 144, 144)
#define ORIENT_FILE L"orient.csv"
#define MAX_ORIENTS 8

class C3DutilsApp : public CWinApp
{
public:
	// enums
	typedef enum {
		OM_BELONGSANY     = 0x0000,
		OM_BELONGSPART    = 0x0001,
		OM_BELONGSBODY    = 0x0002,
		OM_BELONGSMASK    = OM_BELONGSPART | OM_BELONGSBODY,
		OM_USECOLORCUSTOM = 0x0010,
		OM_USECOLOROWNER  = 0x0020,
		OM_USECOLORSOURCE = 0x0040,
		OM_USECOLORMASK   = OM_USECOLORCUSTOM | OM_USECOLOROWNER | OM_USECOLORSOURCE
	} OBJ_FILTER;

	typedef enum {
		CT_EXACT = 0,
		CT_ANY = 765
	} COLOR_TOLERANCE;

	typedef struct orient {
		double angleX;
		double angleY;
		double angleZ;
		WCHAR nameshort[8];
		WCHAR namefull[32];
	} orient_t;

	// init
	C3DutilsApp();
	~C3DutilsApp();

	virtual BOOL InitInstance();
	virtual BOOL ExitInstance();

	// document management
	ksAPI7::IKompasDocument3DPtr GetCurrentDocument();

	// Fases: select same color
	BOOL FaceSelectSameColor( ksAPI7::IKompasDocument3DPtr & doc3d, _variant_t & selection, ULONG &selcount, COLORREF color, ULONG tolerance, ULONG filter, ksAPI7::IModelObjectPtr objptr = NULL );
	BOOL FaceSelectSameColor( ksAPI7::IKompasDocument3DPtr & doc3d, _variant_t & selection, ULONG &selcount, ksAPI7::IPart7Ptr & part, COLORREF color, ULONG tolerance, ULONG filter, ksAPI7::IModelObjectPtr objptr = NULL );
	BOOL SelectDialog( ksAPI7::IKompasDocument3DPtr & doc3d );

	// Faces: change color
	BOOL FaceColorChange( ksAPI7::IKompasDocument3DPtr & doc3d, ULONG &chcount, ULONG usecolor, COLORREF color );
	BOOL ColorChangeDialog( ksAPI7::IKompasDocument3DPtr & doc3d );

	// Selection array
	BOOL SelectObjAdd( _variant_t & selection, ksAPI7::IKompasAPIObjectPtr obj );
	BOOL SelectClear( _variant_t & selection );
	BOOL SelectCommit( ksAPI7::IKompasDocument3DPtr & doc3d, _variant_t & selection );

	// Placement
	BOOL PlacementDialog( ksAPI7::IKompasDocument3DPtr & doc3d );
	const std::vector<orient_t>& GetOrients() const { return m_orients; };

	// Variables
	BOOL VariablesExport( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment = Document::EMBODIMENT_TOP);
	BOOL VariablesImport( ksAPI7::IKompasDocument3DPtr & doc3d, LONG embodiment = Document::EMBODIMENT_TOP);

	// Utility
	static UINT ColorDifference( COLORREF c1, COLORREF c2 );
	static BOOL ObjSetColor( ksAPI7::IModelObjectPtr objptr, ULONG usecolor, COLORREF color );
	static CPalette *CreateDefaultPalete();
	static BOOL SetPlacementEulerDeg( IPlacement *place, double rotXangle, double rotYangle, double rotZangle );
	static BOOL EulerDegfromPlacement( IPlacement *place, double& rotXangle, double& rotYangle, double& rotZangle );
	static void ExtractEulerRadXYZ(double* rotmatrix4x4, double& rotXangle, double& rotYangle, double& rotZangle);
	static void SetEulerDegXYZ(double* rotmatrix4x4, double rotXangle, double rotYangle, double rotZangle);

	BOOL hasAPI() { return (m_kompas != NULL) && (m_newKompasAPI != NULL); };
	CString LoadStr( UINT strID );
	SIZE_T LoadStr( UINT strID, LPWSTR buf, SIZE_T bufsz );
	void EnableTaskAccess( BOOL enable );

private:
	// Methods
	BOOL GetKompas();
	UINT LoadOrients();

	// Members
	ksAPI7::IApplicationPtr m_newKompasAPI;               
	KompasObjectPtr m_kompas;
	_locale_t m_numlocale;  // explicit numeric locale to stricten radix character

	// Selection params dialog
	ULONG m_selfilter;
	BOOL m_selcolorovr;
	COLORREF m_selcolor;
	UINT m_seltolerance;

	// Color change dialog
	ULONG m_chusecolor;
	COLORREF m_chcolor;

	std::vector<orient_t> m_orients;
};

extern C3DutilsApp theApp;

#endif // !defined(AFX_3DUTILS_H__8FC6CC92_1590_41F3_83BF_4EC6CA01F2C8__INCLUDED_)
