#pragma once
#include "afxwin.h"

// ColorChangeDlg

class ColorChangeDlg : public CDialog
{
	DECLARE_DYNAMIC(ColorChangeDlg)

public:
	ColorChangeDlg(CWnd* pParent = NULL);
	ColorChangeDlg(ULONG usecolor, COLORREF color, CWnd* pParent = NULL);
	virtual ~ColorChangeDlg();

	enum { IDD = IDD_COLORCHANGEDLG };

	ULONG GetUseColor() { return m_usecolor; };
	void SetUseColor(ULONG usecolor) { m_usecolor = usecolor; };
	COLORREF GetColor() { return m_color; };
	void SetColor(COLORREF color) { m_color = color; };

protected:
	virtual BOOL OnInitDialog();

	DECLARE_MESSAGE_MAP()

private:
	void InitControls();

	CButton *m_customrd;
	CButton *m_ownerrd;
	CButton *m_sourcerd;
	CMFCColorButton *m_colorbtn;

	CPalette *m_defpalette;

	// in/out
	ULONG m_usecolor;
	COLORREF m_color;

public:
	afx_msg void OnColorChange();
	afx_msg void OnBnClickedCustomrd();
	afx_msg void OnBnClickedOwnerrd();
	afx_msg void OnBnClickedSourcerd();
};
