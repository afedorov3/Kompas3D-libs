#pragma once


// SelectParmDlg dialog

class SelectParmDlg : public CDialogEx
{
	DECLARE_DYNAMIC(SelectParmDlg)

public:
	SelectParmDlg(CWnd* pParent = NULL);   // standard constructor
	SelectParmDlg(ULONG filter, CWnd* pParent = NULL);
	SelectParmDlg(ULONG filter, BOOL hassource, BOOL colorovr, COLORREF m_color, UINT tolerance, CWnd* pParent = NULL);
	virtual ~SelectParmDlg();

	ULONG GetFilter() { return m_filter; };
	void SetFilter(ULONG filter) { m_filter = filter; };
	BOOL GetColorOverride() { return m_colorovr; };
	void SetColorOverride(BOOL colorovr) { m_colorovr = colorovr; };
	COLORREF GetColor() { return m_color; };
	void SetColor(COLORREF color) { m_color = color; };
	UINT GetTolerance() { return m_tolerance; };
	void SetTolerance(UINT tolerance) { m_tolerance = tolerance; };

// Dialog Data
	enum { IDD = IDD_SELECTPARAM };

protected:
	virtual BOOL OnInitDialog();
	//virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

	DECLARE_MESSAGE_MAP()

private:
	void InitControls();

	void HasSourceUpdate();
	void BelongsUpdate();
	void HasColorUpdate();
	void UseColorUpdate(ULONG item = 0);
	void ToleranceUpdate();
	void ToleranceDescrUpdate(INT value);

	CComboBox *m_belongscombo;
	CComboBox *m_hascolorcombo;
	CButton *m_usecolorcustomchk;
	CButton *m_usecolorownerchk;
	CButton *m_usecolorsourcechk;
	CMFCColorButton *m_colorbtn;
	CSliderCtrl *m_toleranceslider;
	CStatic *m_tolerancedescr;

	CPalette *m_defpalette;

	// in
	BOOL m_hassource;

	// in/out
	ULONG m_filter;
	BOOL m_colorovr;
	COLORREF m_color;
	UINT m_tolerance;

public:
	afx_msg void OnCbnSelchangeBelongs();
	afx_msg void OnCbnSelchangeHascolor();
	afx_msg void OnBnClickedUsecolorcustom();
	afx_msg void OnBnClickedUsecolorowner();
	afx_msg void OnBnClickedUsecolorsource();
	afx_msg void OnBnClickedSelmfccolorbutton();
	afx_msg void OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar);
};
