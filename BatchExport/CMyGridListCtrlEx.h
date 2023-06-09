////////////////////////////////////////////////////////////////////////////////
//
// CMyGridListCtrlEx.h
// BatchExport library
// Alex Fedorov, 2018
//
////////////////////////////////////////////////////////////////////////////////

#pragma once

#include "CGridListCtrlEx\CGridListCtrlEx.h"
#include "CGridListCtrlEx\CGridListCtrlGroups.h"
#include "CGridListCtrlEx\CGridColumnTrait.h"
#include "CGridListCtrlEx\CGridColumnTraitCombo.h"
#include "CGridListCtrlEx\CGridColumnTraitEdit.h"

#define WM_CELLCHANGEBASE (WM_USER + 100)
#define WM_EMBODIMENTCHANGE WM_CELLCHANGEBASE + CBatchExportApp::VP_EMBODIMENT
#define WM_DIRCHANGE WM_CELLCHANGEBASE + CBatchExportApp::VP_DIR
#define WM_NAMECHANGE WM_CELLCHANGEBASE + CBatchExportApp::VP_NAME
#define WM_COMMENTADDCHANGE WM_CELLCHANGEBASE + CBatchExportApp::VP_COMMENTADD
#define WM_VARIABLESCHANGE WM_CELLCHANGEBASE + CBatchExportApp::VP_VARIABLES
#define WM_OPTIONSCHANGE WM_CELLCHANGEBASE + CBatchExportApp::VP_OPTIONS

class CMyGridListCtrlEx : public CGridListCtrlEx
{
public:
	virtual bool OnDisplayCellTooltip(int nRow, int nCol, CString& strResult);
	virtual bool OnDisplayCellColor(int nRow, int nCol, COLORREF& textColor, COLORREF& backColor);

	BOOL AddHeaderToolTip(int nCol, LPCTSTR sTip = NULL);

protected:
	virtual void PreSubclassWindow();
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	virtual bool OnEditComplete(int nRow, int nCol, CWnd* pEditor, LV_DISPINFO* pLVDI);

	BOOL RegisterDropTarget();
	UINT CellStatus(int nRow, int nCol);

	CToolTipCtrl m_ToolTip;
};
