#pragma once


// CMySliderCtrl

class CMySliderCtrl : public CSliderCtrl
{
	DECLARE_DYNAMIC(CMySliderCtrl)

public:
	CMySliderCtrl();
	virtual ~CMySliderCtrl();

	BOOL AttachSlider(int nCtlID, CWnd* pParentWnd);
	void SetPos(double pos);
	double GetPos();

	afx_msg void OnMButtonDown(UINT nflags, CPoint Point); 
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
protected:
	DECLARE_MESSAGE_MAP()
};
