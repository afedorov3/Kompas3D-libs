////////////////////////////////////////////////////////////////////////////////
//
// OverwriteDlg.h
// BatchExport library
// Alex Fedorov, 2018
//
////////////////////////////////////////////////////////////////////////////////

#pragma once
#include "OverwriteDlg-rc.h"

#define IDNONE 8
#define IDALL 9

// OverwriteDlg dialog

class OverwriteDlg : public CDialog
{
	DECLARE_DYNAMIC(OverwriteDlg)

public:
	enum OVRENTITY {
		ENTITY_FILE = 1,
		ENTITY_DIR,
		ENTITY_VAR
	};

	OverwriteDlg(LPCWSTR caption, OVRENTITY entity, LPCWSTR detail, CWnd* pParent = NULL);   // standard constructor
	virtual ~OverwriteDlg();

// Dialog Data
	enum { IDD = IDD_OVRDLG };

protected:
	virtual BOOL OnInitDialog();

	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	afx_msg void OnBnClickedAll();
	afx_msg void OnBnClickedYes();
	afx_msg void OnBnClickedNo();
	afx_msg void OnBnClickedNone();
	afx_msg void ClipboardCopyDetail();

	DECLARE_MESSAGE_MAP()

private:
	HACCEL m_hAccelTable;
    LPCWSTR m_caption;
	OVRENTITY m_entity;
	LPCWSTR m_detail;
};
