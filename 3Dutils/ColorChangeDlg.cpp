// ColorChangeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "3Dutils.h"
#include "ColorChangeDlg.h"

// ColorChangeDlg

IMPLEMENT_DYNAMIC(ColorChangeDlg, CDialog)

ColorChangeDlg::ColorChangeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(ColorChangeDlg::IDD, pParent)
{
	InitControls();
	m_defpalette = NULL;
	m_usecolor = useColorOur;
	m_color = DEFCOLOR;
}

ColorChangeDlg::ColorChangeDlg(ULONG usecolor, COLORREF color, CWnd* pParent /*=NULL*/)
	: CDialog(ColorChangeDlg::IDD, pParent)
{
	InitControls();
	m_defpalette = NULL;
	m_usecolor = usecolor;
	m_color = color;
}

ColorChangeDlg::~ColorChangeDlg()
{
	if (m_defpalette != NULL)
		delete m_defpalette;
}

void ColorChangeDlg::InitControls()
{
	m_customrd = NULL;
	m_ownerrd = NULL;
	m_sourcerd = NULL;
	m_colorbtn = NULL;
}

BOOL ColorChangeDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	m_customrd = (CButton*)GetDlgItem(IDC_CUSTOMRD);
	m_ownerrd = (CButton*)GetDlgItem(IDC_OWNERRD);
	m_sourcerd = (CButton*)GetDlgItem(IDC_SOURCERD);
	m_colorbtn = (CMFCColorButton*)GetDlgItem(IDC_CHANGEMFCCOLORBUTTON);
	m_defpalette = C3DutilsApp::CreateDefaultPalete();
	if (m_defpalette != NULL)
		m_colorbtn->SetPalette(m_defpalette);
	m_colorbtn->EnableOtherButton(L"Палитра");

	m_customrd->SetCheck(m_usecolor == useColorOur);
	m_ownerrd->SetCheck(m_usecolor == useColorOwner);
	m_sourcerd->SetCheck(m_usecolor == useColorSource);

	m_colorbtn->SetColor(m_color);
	m_colorbtn->ShowWindow(m_usecolor == useColorOur?SW_SHOW:SW_HIDE);

	GetDlgItem(IDOK)->SetFocus();

	return FALSE;  // return TRUE unless you set the focus to a control
}

BEGIN_MESSAGE_MAP(ColorChangeDlg, CDialog)
	ON_BN_CLICKED(IDC_CHANGEMFCCOLORBUTTON, &ColorChangeDlg::OnColorChange)
	ON_BN_CLICKED(IDC_CUSTOMRD, &ColorChangeDlg::OnBnClickedCustomrd)
	ON_BN_CLICKED(IDC_OWNERRD, &ColorChangeDlg::OnBnClickedOwnerrd)
	ON_BN_CLICKED(IDC_SOURCERD, &ColorChangeDlg::OnBnClickedSourcerd)
END_MESSAGE_MAP()

// ColorChangeDlg message handlers

void ColorChangeDlg::OnColorChange()
{
	m_color = m_colorbtn->GetColor();
}

void ColorChangeDlg::OnBnClickedCustomrd()
{
	m_usecolor = useColorOur;
	m_colorbtn->ShowWindow(SW_SHOW);
}

void ColorChangeDlg::OnBnClickedOwnerrd()
{
	m_usecolor = useColorOwner;
	m_colorbtn->ShowWindow(SW_HIDE);
}

void ColorChangeDlg::OnBnClickedSourcerd()
{
	m_usecolor = useColorSource;
	m_colorbtn->ShowWindow(SW_HIDE);
}
