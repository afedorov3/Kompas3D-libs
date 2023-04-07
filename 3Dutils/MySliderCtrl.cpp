// MySliderCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "3Dutils.h"
#include "MySliderCtrl.h"


// CMySliderCtrl

IMPLEMENT_DYNAMIC(CMySliderCtrl, CSliderCtrl)

CMySliderCtrl::CMySliderCtrl()
{

}

CMySliderCtrl::~CMySliderCtrl()
{
}

//
// AttachSlider
//
// Subclasses an edit with ID nCtrlID and with parent pParentWnd
//
//   Uses CWnd::SubClassDlgItem which lets us insert
//   our message map into the MFC message routing system
//
BOOL CMySliderCtrl::AttachSlider(int nCtlID, CWnd* pParentWnd)
{
	ASSERT(pParentWnd != NULL);

	if (SubclassDlgItem(nCtlID, pParentWnd)) {
		ASSERT_VALID(this);
		return TRUE;
	}

	// SubClassDlgItem failed
	TRACE2("CMySliderCtrl, could not attach to %d %x\n", nCtlID, pParentWnd);
	ASSERT(FALSE);
	return FALSE;  
} 

void CMySliderCtrl::SetPos(double pos)
{
	CSliderCtrl::SetPos(static_cast<int>(pos + (pos<0?-0.5:0.5)));
}

double CMySliderCtrl::GetPos()
{
	return static_cast<double>(CSliderCtrl::GetPos());
}

BEGIN_MESSAGE_MAP(CMySliderCtrl, CSliderCtrl)
	ON_WM_MBUTTONDOWN()
	ON_WM_KEYDOWN()
END_MESSAGE_MAP()



// CMySliderCtrl message handlers
void CMySliderCtrl::OnMButtonDown(UINT nflags, CPoint Point)
{
	SetPos(0.0);
	this->GetParent()->PostMessage(WM_HSCROLL, (WPARAM)SB_THUMBTRACK, (LPARAM)this->m_hWnd);
}

void CMySliderCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if ((nChar == VK_INSERT) || (nChar == VK_DELETE)) {
		SetPos(0.0);
		this->GetParent()->PostMessage(WM_HSCROLL, (WPARAM)SB_THUMBTRACK, (LPARAM)this->m_hWnd);
		return;
	}
	CSliderCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
}
