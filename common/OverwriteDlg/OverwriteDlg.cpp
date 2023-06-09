////////////////////////////////////////////////////////////////////////////////
//
// OverwriteDlg.cpp
// BatchExport library
// Alex Fedorov, 2018
//
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "OverwriteDlg.h"
#include "afxdialogex.h"
#include "../utils.h" // Str2Clipboard()

// OverwriteDlg dialog

IMPLEMENT_DYNAMIC(OverwriteDlg, CDialog)

OverwriteDlg::OverwriteDlg(LPCWSTR caption, OVRENTITY entity, LPCWSTR detail, CWnd* pParent /*=NULL*/)
	: CDialog(OverwriteDlg::IDD, pParent)
{
	m_caption = caption;
	m_entity = entity;
	m_detail = detail;
}

OverwriteDlg::~OverwriteDlg()
{
}

void OverwriteDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BOOL OverwriteDlg::PreTranslateMessage(MSG* pMsg) 
  {
	if ( TranslateAccelerator( m_hWnd, m_hAccelTable, pMsg ) ) return TRUE;

	return CDialog::PreTranslateMessage(pMsg);
}

BEGIN_MESSAGE_MAP(OverwriteDlg, CDialog)
	ON_BN_CLICKED(IDC_OVRALL, OnBnClickedAll)
	ON_BN_CLICKED(IDC_OVRYES, OnBnClickedYes)
	ON_BN_CLICKED(IDC_OVRNO, OnBnClickedNo)
	ON_BN_CLICKED(IDC_OVRNONE, OnBnClickedNone)
	ON_COMMAND(ID_OVRACCELERATOR_CTRLINS, ClipboardCopyDetail)
	ON_COMMAND(ID_OVRACCELERATOR_CTRLC, ClipboardCopyDetail)
	ON_COMMAND(ID_OVRACCELERATOR_Y, OnBnClickedYes)
	ON_COMMAND(ID_OVRACCELERATOR_YR, OnBnClickedYes)
	ON_COMMAND(ID_OVRACCELERATOR_N, OnBnClickedNo)
	ON_COMMAND(ID_OVRACCELERATOR_NR, OnBnClickedNo)
	ON_COMMAND(ID_OVRACCELERATOR_A, OnBnClickedAll)
	ON_COMMAND(ID_OVRACCELERATOR_AR, OnBnClickedAll)
	ON_COMMAND(ID_OVRACCELERATOR_S, OnBnClickedNone)
	ON_COMMAND(ID_OVRACCELERATOR_SR, OnBnClickedNone)
END_MESSAGE_MAP()

BOOL OverwriteDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	SetWindowText(m_caption);
	CStatic *en = (CStatic*)GetDlgItem(IDC_OVRENTITY);
	CStatic *det = (CStatic*)GetDlgItem(IDC_OVRDETAIL);
	switch(m_entity) {
	case ENTITY_FILE:
		en->SetWindowText(L"Файл");
		det->ModifyStyle(SS_ENDELLIPSIS|SS_WORDELLIPSIS, SS_PATHELLIPSIS);
		break;
	case ENTITY_DIR:
		en->SetWindowText(L"Каталог");
		det->ModifyStyle(SS_ENDELLIPSIS|SS_WORDELLIPSIS, SS_PATHELLIPSIS);
		break;
	case ENTITY_VAR:
		en->SetWindowText(L"Переменная");
		det->ModifyStyle(SS_ENDELLIPSIS|SS_PATHELLIPSIS, SS_WORDELLIPSIS);
		break;
	}

	GetDlgItem(IDC_OVRDETAIL)->SetWindowText(m_detail);
	m_hAccelTable = LoadAccelerators(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_OVRACCELERATORS));

	return TRUE;
}

void OverwriteDlg::OnBnClickedAll()
{
	EndDialog(IDALL);
}

void OverwriteDlg::OnBnClickedYes()
{
	EndDialog(IDYES);
}

void OverwriteDlg::OnBnClickedNo()
{
	EndDialog(IDNO);
}

void OverwriteDlg::OnBnClickedNone()
{
	EndDialog(IDNONE);
}

void OverwriteDlg::ClipboardCopyDetail()
{
	Utils::Str2Clipboard(m_detail);
}
