// SelectParmDlg.cpp : implementation file
//

#include "stdafx.h"
#include "3Dutils.h"
#include "SelectParmDlg.h"
#include "afxdialogex.h"

#include <string>

/*static inline LONG ToLong( const CString& str ) {
	return std::stol( str.GetString() );
}*/

// SelectParmDlg dialog

IMPLEMENT_DYNAMIC(SelectParmDlg, CDialogEx)

/*void SelectParmDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}*/

SelectParmDlg::SelectParmDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(SelectParmDlg::IDD, pParent)
{
	InitControls();
	m_defpalette = NULL;
	m_filter = C3DutilsApp::OM_BELONGSANY | C3DutilsApp::OM_USECOLORMASK;
	m_hassource = FALSE;
	m_colorovr = TRUE;
	m_color = DEFCOLOR;
	m_tolerance = 0;
}

SelectParmDlg::SelectParmDlg(ULONG filter, CWnd* pParent /*=NULL*/)
	: CDialogEx(SelectParmDlg::IDD, pParent)
{
	InitControls();
	m_defpalette = NULL;
	m_filter = filter;
	m_hassource = FALSE;
	m_colorovr = TRUE;
	m_color = DEFCOLOR;
	m_tolerance = 0;
}

SelectParmDlg::SelectParmDlg(ULONG filter, BOOL hassource, BOOL colorovr, COLORREF color, 
							 UINT tolerance, CWnd* pParent /*=NULL*/)
	: CDialogEx(SelectParmDlg::IDD, pParent)
{
	InitControls();
	m_defpalette = NULL;
	m_filter = filter;
	m_hassource = hassource;
	m_colorovr = colorovr;
	m_color = color;
	m_tolerance = tolerance;
}

SelectParmDlg::~SelectParmDlg()
{
	if (m_defpalette != NULL)
		delete m_defpalette;
}

BOOL SelectParmDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_belongscombo = (CComboBox*)GetDlgItem(IDC_BELONGS);
	m_hascolorcombo = (CComboBox*)GetDlgItem(IDC_HASCOLOR);
	m_usecolorcustomchk = (CButton*)GetDlgItem(IDC_USECOLORCUSTOM);
	m_usecolorownerchk = (CButton*)GetDlgItem(IDC_USECOLOROWNER);
	m_usecolorsourcechk = (CButton*)GetDlgItem(IDC_USECOLORSOURCE);
	m_colorbtn = (CMFCColorButton*)GetDlgItem(IDC_SELMFCCOLORBUTTON);
	m_defpalette = C3DutilsApp::CreateDefaultPalete();
	if (m_defpalette != NULL)
		m_colorbtn->SetPalette(m_defpalette);
	m_colorbtn->EnableOtherButton(L"Палитра");
	m_toleranceslider = (CSliderCtrl*)GetDlgItem(IDC_TOLERANCESLIDER);
	m_toleranceslider->SetRange(0, 100);
	m_toleranceslider->SetPageSize(10);
	m_toleranceslider->SetTicFreq(10);
	m_tolerancedescr = (CStatic*)GetDlgItem(IDC_TOLERANCEDESCR);

	HasSourceUpdate();
	BelongsUpdate();
	HasColorUpdate();
	UseColorUpdate();
	ToleranceUpdate();

	GetDlgItem(IDOK)->SetFocus();

	return FALSE;
}

void SelectParmDlg::InitControls()
{
	m_belongscombo = NULL;
	m_hascolorcombo = NULL;
	m_usecolorcustomchk = NULL;
	m_usecolorownerchk = NULL;
	m_usecolorsourcechk = NULL;
	m_colorbtn = NULL;
	m_toleranceslider = NULL;
	m_tolerancedescr = NULL;
}

BEGIN_MESSAGE_MAP(SelectParmDlg, CDialogEx)
	ON_CBN_SELCHANGE(IDC_BELONGS, &SelectParmDlg::OnCbnSelchangeBelongs)
	ON_CBN_SELCHANGE(IDC_HASCOLOR, &SelectParmDlg::OnCbnSelchangeHascolor)
	ON_BN_CLICKED(IDC_USECOLORCUSTOM, &SelectParmDlg::OnBnClickedUsecolorcustom)
	ON_BN_CLICKED(IDC_USECOLOROWNER, &SelectParmDlg::OnBnClickedUsecolorowner)
	ON_BN_CLICKED(IDC_USECOLORSOURCE, &SelectParmDlg::OnBnClickedUsecolorsource)
	ON_BN_CLICKED(IDC_SELMFCCOLORBUTTON, &SelectParmDlg::OnBnClickedSelmfccolorbutton)
	ON_WM_HSCROLL()
END_MESSAGE_MAP()

// SelectParmDlg message handlers

void SelectParmDlg::HasSourceUpdate()
{
	if (!m_hassource) {
		m_belongscombo->SetCurSel(0);
		m_hascolorcombo->SetCurSel(1);
	}
	m_belongscombo->EnableWindow(m_hassource > 0);
	m_hascolorcombo->EnableWindow(m_hassource > 0);
}

void SelectParmDlg::BelongsUpdate()
{
	if (!m_hassource)
		return;

	switch(m_filter & C3DutilsApp::OM_BELONGSMASK) {
	case C3DutilsApp::OM_BELONGSPART: m_belongscombo->SetCurSel(1); break;
	case C3DutilsApp::OM_BELONGSBODY: m_belongscombo->SetCurSel(2); break;
	default: m_belongscombo->SetCurSel(0); break;
	}
}

void SelectParmDlg::HasColorUpdate()
{
	if (m_hassource)
		m_hascolorcombo->SetCurSel(m_colorovr?1:0);

	m_colorbtn->ShowWindow(m_hascolorcombo->GetCurSel() != 0?SW_SHOW:SW_HIDE);
	m_colorbtn->SetColor(m_color);
}

void SelectParmDlg::UseColorUpdate(ULONG item)
{
	if (item != 0) {
		if ((m_filter & C3DutilsApp::OM_USECOLORMASK) != item) // enforce at least one option
			m_filter ^= item;
	}

	m_usecolorcustomchk->SetCheck(m_filter & C3DutilsApp::OM_USECOLORCUSTOM);
	m_usecolorownerchk->SetCheck(m_filter & C3DutilsApp::OM_USECOLOROWNER);
	m_usecolorsourcechk->SetCheck(m_filter & C3DutilsApp::OM_USECOLORSOURCE);
}

void SelectParmDlg::ToleranceUpdate()
{
	INT pos = (INT)((double)(m_tolerance * 100) / C3DutilsApp::CT_ANY + 0.5);
	m_toleranceslider->SetPos(pos);
	ToleranceDescrUpdate(pos);
}

void SelectParmDlg::ToleranceDescrUpdate(INT value)
{
	CString descr(L"Допуск совпадения цвета");

	switch(value) {
	case 0: descr += L" (Точный)"; break;
	case 100: descr += L" (Любой)"; break;
	default:
		descr.AppendFormat(L" (%d%%)", value);
	}
	m_tolerancedescr->SetWindowText(descr);
}

void SelectParmDlg::OnCbnSelchangeBelongs()
{
	m_filter &= ~C3DutilsApp::OM_BELONGSMASK;
	switch(m_belongscombo->GetCurSel()) {
	case 1: m_filter |= C3DutilsApp::OM_BELONGSPART; break;
	case 2: m_filter |= C3DutilsApp::OM_BELONGSBODY; break;
	default: m_filter |= C3DutilsApp::OM_BELONGSANY; break;
	}
}

void SelectParmDlg::OnCbnSelchangeHascolor()
{
	if (m_hascolorcombo->GetCurSel() != 0)
		m_colorovr = TRUE;
	else
		m_colorovr = FALSE;
	m_colorbtn->ShowWindow(m_colorovr?SW_SHOW:SW_HIDE);
}

void SelectParmDlg::OnBnClickedUsecolorcustom()
{
	UseColorUpdate(C3DutilsApp::OM_USECOLORCUSTOM);
}

void SelectParmDlg::OnBnClickedUsecolorowner()
{
	UseColorUpdate(C3DutilsApp::OM_USECOLOROWNER);
}

void SelectParmDlg::OnBnClickedUsecolorsource()
{
	UseColorUpdate(C3DutilsApp::OM_USECOLORSOURCE);
}

void SelectParmDlg::OnBnClickedSelmfccolorbutton()
{
	m_color = m_colorbtn->GetColor();
}

void SelectParmDlg::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	if (pScrollBar == (CWnd*)m_toleranceslider) {
		INT pos = m_toleranceslider->GetPos();
		m_tolerance = (UINT)((double)(pos * C3DutilsApp::CT_ANY) / 100.0 + 0.5);
		ToleranceDescrUpdate(pos);
	}

	CDialogEx::OnHScroll(nSBCode, nPos, pScrollBar);
}
