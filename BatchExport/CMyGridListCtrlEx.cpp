////////////////////////////////////////////////////////////////////////////////
//
// CMyGridListCtrlEx.cpp
// BatchExport library
// Alex Fedorov, 2018
//
// Uses ListCtrl header ToolTip code kindly provided by
// Zafir Anjum @ https://www.codeguru.com/cpp/controls/listview/tooltiptitletip/article.php/c993/Tooltip-for-individual-column-header.htm
//
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CMyGridListCtrlEx.h"
#include "GridViewDlg.h"
#include "BatchExport.h"

#pragma warning(disable:4100)	// unreferenced formal parameter

void CMyGridListCtrlEx::PreSubclassWindow()
{
    CGridListCtrlEx::PreSubclassWindow();
    m_ToolTip.Create( this );
}

BOOL CMyGridListCtrlEx::PreTranslateMessage(MSG* pMsg) 
{
	m_ToolTip.RelayEvent( pMsg );	
	return CGridListCtrlEx::PreTranslateMessage(pMsg);
}

bool CMyGridListCtrlEx::OnDisplayCellTooltip(int nRow, int nCol, CString& strResult)
{
	BOOL ret = TRUE;
	UINT status = CellStatus(nRow, nCol);
	if ((status != LIBSTATUS_SUCCESS) &&
		(status <= LIBSTATUS_USER_ERR_MAX)) {
		BOOL ret = CBatchExportApp::StrStatus(strResult, status);
		return (ret?true:false); // convert to bool
	}
	strResult = GetItemText(nRow, nCol);
	return true;
}

bool CMyGridListCtrlEx::OnDisplayCellColor(int nRow, int nCol, COLORREF& textColor, COLORREF& backColor)
{
	UINT status = CellStatus(nRow, nCol);
	if ((status != LIBSTATUS_SUCCESS) &&
		((status & LIBSTATUS_EXCLUSIVE) != LIBSTATUS_INFO)) {
		textColor = RGB(250,75,85);
		return true;
	}

	return false;  // Use default color
}

bool CMyGridListCtrlEx::OnEditComplete(int nRow, int nCol, CWnd* pEditor, LV_DISPINFO* pLVDI)
{
	CGridColumnTrait* pTrait = GetCellColumnTrait(nRow, nCol);
	if (pTrait != NULL)
		pTrait->OnEditEnd();
	if (pLVDI == NULL)
		return false;	// Parent view rejected LVN_BEGINLABELEDIT notification

	bool txtEdit = pLVDI->item.mask & LVIF_TEXT && pLVDI->item.pszText != NULL;
	bool imgEdit = pLVDI->item.mask & LVIF_IMAGE && pLVDI->item.iImage != I_IMAGECALLBACK;

	if (txtEdit) {	// Update data
		LRESULT res = GetParent()->SendMessage(WM_CELLCHANGEBASE + nCol - 1, (WPARAM)nRow, (LPARAM)pLVDI->item.pszText);
		if (res == 0)
			txtEdit = false;
	}
	return txtEdit || imgEdit;	// Accept edit
}

// disable Drag&Drop
BOOL CMyGridListCtrlEx::RegisterDropTarget()
{
	return FALSE;
}

UINT CMyGridListCtrlEx::CellStatus(int nRow, int nCol)
{
	if (nRow < 0)
		return LIBSTATUS_ERR_INVARG;
	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)GetItemData(nRow);
	if (varptr == NULL)
		return LIBSTATUS_ERR_INVARG;
	ULONG status = varptr->status;
	// exclusive errors
	if ((status == LIBSTATUS_SUCCESS) ||
		((status & LIBSTATUS_EXCLUSIVE) != LIBSTATUS_SUCCESS))
		return status;
	// non exclusive errors
	switch(nCol) {
	case CBatchExportApp::VP_EMBODIMENT + 1: status &= LIBSTATUS_EMBODIMENT_MASK; break;
	case CBatchExportApp::VP_DIR + 1:        status &= LIBSTATUS_DIR_MASK;	break;
	case CBatchExportApp::VP_NAME + 1:       status &= LIBSTATUS_NAME_MASK; break;
	case CBatchExportApp::VP_COMMENTADD + 1: status &= LIBSTATUS_COMMENTADD_MASK; break;
	case CBatchExportApp::VP_VARIABLES + 1:  status &= LIBSTATUS_VARS_MASK; break;
	case CBatchExportApp::VP_OPTIONS + 1:    status &= LIBSTATUS_OPTS_MASK; break;
	default:
		status = LIBSTATUS_SUCCESS;
	}

	return status;
}

// by Zafir Anjum
// https://www.codeguru.com/cpp/controls/listview/tooltiptitletip/article.php/c993/Tooltip-for-individual-column-header.htm
BOOL CMyGridListCtrlEx::AddHeaderToolTip(int nCol, LPCTSTR sTip)
{
	const int TOOLTIP_LENGTH = 80;
	WCHAR buf[TOOLTIP_LENGTH+1];

	CHeaderCtrl* pHeader = (CHeaderCtrl*)GetDlgItem(0);
	int nColumnCount = pHeader->GetItemCount();
	if( nCol >= nColumnCount)
		return FALSE;

	if( (GetStyle() & LVS_TYPEMASK) != LVS_REPORT )
		return FALSE;

	// Get the header height
	RECT rect;
	pHeader->GetClientRect( &rect );
	int height = rect.bottom;

	RECT rctooltip;
	rctooltip.top = 0;
	rctooltip.bottom = rect.bottom;

	// Now get the left and right border of the column
	rctooltip.left = 0 - GetScrollPos( SB_HORZ );
	for( int i = 0; i < nCol; i++ )
		rctooltip.left += GetColumnWidth( i );
	rctooltip.right = rctooltip.left + GetColumnWidth( nCol );

	if( sTip == NULL )
	{
		// Get column heading
		LV_COLUMN lvcolumn;
		lvcolumn.mask = LVCF_TEXT;
		lvcolumn.pszText = buf;
		lvcolumn.cchTextMax = TOOLTIP_LENGTH;
		if( !GetColumn( nCol, &lvcolumn ) )
			return FALSE;
	}

	m_ToolTip.AddTool( GetDlgItem(0), sTip ? sTip : buf, &rctooltip, nCol+1 );
	return TRUE;
}
