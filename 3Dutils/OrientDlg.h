#pragma once
#include "NumericEdit\NumericEdit.h"
#include "MySliderCtrl.h"

// OrientDlg dialog

class OrientDlg : public CDialogEx
{
	DECLARE_DYNAMIC(OrientDlg)

public:
	typedef enum {
		AUTOPLACE_DISABLED = 0,
		// Don't change order of the folowing
		AUTOPLACE_TOPLEFT,
		AUTOPLACE_TOPRIGHT,
		AUTOPLACE_BOTLEFT,
		AUTOPLACE_BOTRIGHT,
		AUTOPLACE_DEFAULT = AUTOPLACE_BOTRIGHT
	} AUTOPLACE;

	OrientDlg(ksAPI7::IKompasDocument3DPtr& doc3d, LONG posX = LONG_MIN, LONG posY = LONG_MIN, bool preview = false, CWnd* pParent = NULL);
	virtual ~OrientDlg();

	double GetX() { return m_curX; };
	double GetY() { return m_curY; };
	double GetZ() { return m_curZ; };
	bool GetPreview() { return m_preview; };
	void SetPreview(bool enabled) { m_preview = enabled; };
	void GetWinPos(LONG& PosX, LONG& PosY) { PosX = m_wPosX; PosY = m_wPosY; };
	void SetWinPos(LONG PosX, LONG PosY) { m_wPosX = PosX; m_wPosY = PosY; };
	UINT GetAutoplace() { return (UINT)m_autoplace; };
	void SetAutoplace(UINT autoplace) { m_autoplace = autoplace > AUTOPLACE_BOTRIGHT ? AUTOPLACE_BOTRIGHT : (AUTOPLACE)autoplace; };

// Dialog Data
	enum { IDD = IDD_ORIENT };

private:
	enum {
		CTRLX = 0x1,
		CTRLY = 0x2,
		CTRLZ = 0x4,
		ALL = CTRLX|CTRLY|CTRLZ
	};

	enum {
		PREVIEW_TIMER = 1
	};

	void AddOrientBtns();
	void PlaceWindow(LONG height = LONG_MIN);
	void UpdateControls(int sel = ALL);
	void Orient();
	void Preview(UINT delay = 0);
	BOOL AddDocProjection(const C3DutilsApp::orient_t& orient);

	IViewProjectionCollectionPtr m_vpcol;
	IViewProjectionPtr m_vproj;
	IPlacementPtr m_place;
	HWND m_drawHwnd;
	double m_initX, m_initY, m_initZ;
	double m_curX, m_curY, m_curZ;
	bool m_preview;
	bool m_active;
	bool m_tooltip;
	LONG m_wPosX, m_wPosY, m_wWidth, m_wHeight, m_wHeightExt, m_wCaptHeight;
	UINT_PTR m_previewTimer;
	AUTOPLACE m_autoplace;
	const std::vector<C3DutilsApp::orient_t>& m_orients;
	std::vector<CButton*> m_oribtns;

	CNumericEdit m_angleX;
	CNumericEdit m_angleY;
	CNumericEdit m_angleZ;
	CMySliderCtrl m_sliderX;
	CMySliderCtrl m_sliderY;
	CMySliderCtrl m_sliderZ;
	CButton *m_previewBtn;

	CButton *m_extBtn;
	CButton *m_addBtn;
	CStatic *m_oriSplit;
	bool m_extended;
	bool m_addtodoc;
	HICON m_hExtDn;
	HICON m_hExtUp;
	HICON m_hDocArrOff;
	HICON m_hDocArrOn;

	HACCEL m_hAccelTable;

protected:
	CToolTipCtrl m_toolTip;

	virtual BOOL OnInitDialog();

	//virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnDestroy();
	afx_msg BOOL OnNcActivate(BOOL);
	afx_msg void OnClose();
	afx_msg void OnHScroll(UINT, UINT, CScrollBar*);
	afx_msg void OnEnKillfocusEditX();
	afx_msg void OnEnKillfocusEditY();
	afx_msg void OnEnKillfocusEditZ();
	afx_msg void OnNotifyEditX(NMHDR*, LRESULT*);
	afx_msg void OnNotifyEditY(NMHDR*, LRESULT*);
	afx_msg void OnNotifyEditZ(NMHDR*, LRESULT*);
	afx_msg void OnStnClickedIconX();
	afx_msg void OnStnClickedIconY();
	afx_msg void OnStnClickedIconZ();
	afx_msg void OnBnClickedReset();
	afx_msg void OnBnClickedPreview();
	afx_msg void OnBnClickedExtBtn();
	afx_msg void OnBnClickedAddBtn();
	afx_msg void OnBnClickedOrient(UINT);
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnTimer(UINT_PTR);
	afx_msg void OnMoving(UINT, LPRECT);
	afx_msg void OnNcLButtonDblClk(UINT, CPoint);
};
