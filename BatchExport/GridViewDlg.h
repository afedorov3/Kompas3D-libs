////////////////////////////////////////////////////////////////////////////////
//
// GridViewDlg.h
// BatchExport library
// Alex Fedorov, 2018
//
////////////////////////////////////////////////////////////////////////////////

#pragma once
#include "afxcmn.h"
#include "resource.h"

#include "CMyGridListCtrlEx.h"
#include "BatchExport.h"

// GridViewDlg dialog
class GridViewDlg : public CDialog
{
// Construction
public:
	typedef enum {
		// behavior options
		GVOPT_READONLY		= 0x1,		// changing variations is not allowed
		GVOPT_SKIPEXCL		= 0x2,		// do not show excluded variations
		GVOPT_PRESERVETEMPL	= 0x4,		// preserve text fields as is, irrelevant if GVOPT_READONLY is set
		GVOPT_ALLOWSAVE		= 0x8,		// allow saving of changes, has a priority over the back button, irrelevant if GVOPT_READONLY is set
		GVOPT_CANBACK		= 0x10,		// show back button inplace of the save button
		GVOPT_ERRIGNORE		= 0x20,		// ignore errors
		GVOPT_NOCOMPINC		= 0x40,		// Include checkbox affects only options, no composition of optstr is done on include flag change
		// layout options
		GVOPT_SHOWCANCEL	= 0x10000,	// show cancel button
		GVOPT_SHOWFILEOPTS	= 0x20000,	// show file related options
		GVOPT_SHOWINCLUDE	= 0x40000,	// show include column
		GVOPT_SHOWVARIABLES	= 0x80000,	// show variables column
		GVOPT_SHOWSTATUS	= 0x100000,	// show status column
		// default options
		GVOPT_DEFAULT = GVOPT_READONLY | GVOPT_PRESERVETEMPL | GVOPT_SHOWCANCEL | GVOPT_SHOWINCLUDE | GVOPT_SHOWVARIABLES
	} OPTIONS;
	typedef enum {
		NO_GROUP = 0,
		GROUP_BY_ATTR = 1,
		GROUP_BY_ATTRROW = 3
	} GROUP_BY;
	GridViewDlg(LPCWSTR title, LPCWSTR OKBtnText, ksAPI7::IKompasDocument3DPtr doc3d,
		CBatchExportApp::VARIATIONS * variations, OPTIONS options = GVOPT_DEFAULT, GROUP_BY groupby = NO_GROUP, CWnd * pParent = NULL);

	void SetOptions(OPTIONS options) { m_options = options; };
	OPTIONS GetOptions() { return (OPTIONS)m_options; };
	void EnableOptions(OPTIONS options) { m_options |= (ULONG)options; };
	void DisableOptions(OPTIONS options) { m_options &= ~(ULONG)options; };
	inline bool HasOpts(OPTIONS options) { return (m_options & options) == options; };	// all of given options are enabled
	inline bool NoOpts(OPTIONS options) { return (m_options & options) == 0; };			// none of given options are enabled

	GROUP_BY GetGroupBy() { return m_groupby; };
	void SetGroupBy(GROUP_BY groupby) { m_groupby = groupby; };

	UINT GetOverwriteFiles() { return m_overwritefiles; };
	void SetOverWriteFiles(UINT overwrite) { m_overwritefiles = overwrite; };
	UINT GetTestMode() { return m_testmode; };
	void SetTestMode(UINT testmode) { m_testmode = testmode; };

// Dialog Data
	enum { IDD = IDD_GRIDVIEW_DIALOG };

// Implementation
protected:
	HICON m_hIcon;
	CToolTipCtrl m_ToolTip;

	virtual BOOL OnInitDialog();

	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	// Generated message map functions
	afx_msg void OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI);
	afx_msg void OnSize( UINT nType, int cx, int cy );
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnBnClickedSave();
	afx_msg void OnBnClickedBack();
	afx_msg void OnBnClickedOverwrite();
	afx_msg void OnBnClickedTestMode();
	afx_msg void OnHeaderClick(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnLvnItemchangedList(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg LRESULT OnEmbodimentChange(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnDirChange(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnNameChange(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnCommentaddChange(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnVariablesChange(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnOptionsChange(WPARAM wParam, LPARAM lParam);
	afx_msg void ClipboardCopy();

	DECLARE_MESSAGE_MAP()

private:
	CMyGridListCtrlEx m_ListCtrl;
	CButton *m_OKbtn;
	CButton *m_Cancelbtn;
	CButton *m_Savebtn;
	CButton *m_Backbtn;
	CButton *m_Overwritechk;
	CButton *m_TestModechk;
	HACCEL m_hAccelTable;
	ksAPI7::IKompasDocument3DPtr m_doc3d;
	CBatchExportApp::VARIATIONS *m_variations;
	CString m_title;
	CString m_okbtntext;

	ULONG m_options;
	GROUP_BY m_groupby;
	UINT m_overwritefiles;
	UINT m_testmode;

	BOOL ItemInclude(UINT nRow, BOOL include);
	void MoveControls(CRect *rc);
	void AllowSave();
	void ForbidSave();
	void EnableOK(UINT enable);
};
