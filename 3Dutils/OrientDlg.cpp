// OrientDlg.cpp : implementation file
//

#include "stdafx.h"
#include "afxdialogex.h"

#include "3Dutils.h"
#include "OrientDlg.h"

#include <strsafe.h>
#include <math.h>
#include <vector>

// OrientDlg dialog

IMPLEMENT_DYNAMIC(OrientDlg, CDialogEx)

OrientDlg::OrientDlg(ksAPI7::IKompasDocument3DPtr& doc3d, BOOL preview /* = FALSE */, UINT autoplace /* = AUTOPLACE_DEFAULT*/,
                     LONG posX /* = LONG_MIN */, LONG posY /*= LONG_MIN*/, CWnd* pParent /* = NULL*/)
	: CDialogEx(OrientDlg::IDD, pParent),
	m_vproj(NULL),
	m_place(NULL),
	m_drawHwnd(NULL),
	m_previewBtn(NULL),
	m_initX(0.0),
	m_initY(0.0),
	m_initZ(0.0),
	m_curX(0.0),
	m_curY(0.0),
	m_curZ(0.0),
	m_active(FALSE),
	m_tooltip(FALSE),
	m_previewTimer(0),
	m_preview(preview),
	m_wPosX(posX),
	m_wPosY(posY),
	m_autoplace(autoplace),
	m_orients(theApp.GetOrients()),
	m_extBtn(NULL),
	m_addBtn(NULL),
	m_extended(FALSE),
	m_addtodoc(FALSE),
	m_hExtDn(NULL),
	m_hExtUp(NULL),
	m_hDocArrOff(NULL),
	m_hDocArrOn(NULL)
{
	if (doc3d == NULL)
		return;

	// document 3D API7 to API5
	IDocument3DPtr doc3d5( IUnknownPtr(ksTransferInterface( doc3d, ksAPI3DCom, 0/*any document*/ ), false/*don't AddRef*/) );
	if (doc3d5 == NULL)
		return;

	m_vpcol = doc3d5->GetViewProjectionCollection();
	if (m_vpcol == NULL)
		return;

	m_vproj = m_vpcol->NewViewProjection();
	if (m_vproj == NULL) {
		return;
	}

	m_place = m_vproj->GetPlacement();

	ksAPI7::IDocumentFramesPtr docFrames( doc3d->GetDocumentFrames());
	if (docFrames == NULL)
		return;

	ksAPI7::IDocumentFramePtr docFrame( docFrames->GetItem( _variant_t( (long)0 ) ) );
	if (docFrame == NULL)
		return;

	m_drawHwnd = (HWND)docFrame->GetHWND();
}

OrientDlg::~OrientDlg()
{
	if (m_place)
		m_place->Release();
	if (m_vproj)
		m_vproj->Release();
	if (m_vpcol)
		m_vpcol->Release();
}

#define ZEROVAL(v) do if ((v <= 0.0000000001) && (v >= -0.0000000001)) v = 0.0; while(0)
BOOL OrientDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	m_angleX.AttachEdit(IDC_EDITX, this);
	m_angleY.AttachEdit(IDC_EDITY, this);
	m_angleZ.AttachEdit(IDC_EDITZ, this);
	m_sliderX.AttachSlider(IDC_SLIDERX, this);
	m_sliderY.AttachSlider(IDC_SLIDERY, this);
	m_sliderZ.AttachSlider(IDC_SLIDERZ, this);
	m_previewBtn = (CButton*)GetDlgItem(IDC_PREVIEW);

	if (m_place)
		C3DutilsApp::EulerDegfromPlacement(m_place, m_initX, m_initY, m_initZ);
	ZEROVAL(m_initX);
	ZEROVAL(m_initY);
	ZEROVAL(m_initZ);
	m_curX = m_initX;
	m_curY = m_initY;
	m_curZ = m_initZ;

	//sliderX->SetBuddy(angleX);
	m_sliderX.SetRange(-180, 180);
	m_sliderX.SetTicFreq(15);
	m_sliderX.SetPageSize(15);

	//sliderY->SetBuddy(angleY);
	m_sliderY.SetRange(-180, 180);
	m_sliderY.SetTicFreq(15);
	m_sliderY.SetPageSize(15);

	//sliderZ->SetBuddy(angleZ);
	m_sliderZ.SetRange(-180, 180);
	m_sliderZ.SetTicFreq(15);
	m_sliderZ.SetPageSize(15);

	m_angleX.SetMinValue(-180.0);
	m_angleX.SetMaxValue(180.0);

	m_angleY.SetMinValue(-180.0);
	m_angleY.SetMaxValue(180.0);

	m_angleZ.SetMinValue(-180.0);
	m_angleZ.SetMaxValue(180.0);

	m_previewBtn->SetCheck(m_preview?BST_CHECKED:BST_UNCHECKED);

	UpdateControls();

	if (m_toolTip.Create(this)) {
		m_tooltip = TRUE;
		m_toolTip.AddTool(GetDlgItem(IDC_ICONX), L"Вокруг оси X");
		m_toolTip.AddTool(GetDlgItem(IDC_ICONY), L"Вокруг оси Y");
		m_toolTip.AddTool(GetDlgItem(IDC_ICONZ), L"Вокруг оси Z");
		m_toolTip.AddTool(GetDlgItem(IDC_PREVIEW), L"Применять ориентацию при изменении углов");
		m_toolTip.AddTool(GetDlgItem(IDC_RESET), L"Вернуть первоначальную ориентацию");
	}

	RECT rect = {};
	GetWindowRect(&rect);
	m_wWidth = rect.right - rect.left;
	m_wHeight = m_wHeightExt = rect.bottom - rect.top;
	GetClientRect(&rect);
	m_wCaptHeight = m_wHeight - rect.bottom;

	AddOrientBtns(); // uses m_wWidth, m_wHeight, m_toolTip; sets m_wHeightExt
	PlaceWindow(); // uses m_wWidth, m_wHeight

	m_hAccelTable = LoadAccelerators(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_ORIENTACCELERATORS));

	m_active = TRUE;

	return FALSE;
}

void OrientDlg::Confirm()
{
	m_active = FALSE;

	if (m_previewTimer != 0)
		KillTimer(m_previewTimer), m_previewTimer = 0;

	// update data in case of NC activate
	switch(::GetDlgCtrlID(::GetFocus())) {
	case IDC_EDITX:
		m_angleX.GetData(m_curX);
		break;
	case IDC_EDITY:
		m_angleY.GetData(m_curY);
		break;
	case IDC_EDITZ:
		m_angleZ.GetData(m_curZ);
		break;
	}

	Orient();

	EndDialog(IDOK);
}

void OrientDlg::Cancel()
{
	m_active = FALSE;

	if (m_previewTimer != 0)
		KillTimer(m_previewTimer), m_previewTimer = 0;

	m_curX = m_initX;
	m_curY = m_initY;
	m_curZ = m_initZ;
	UpdateControls();

	Orient();

	EndDialog(IDCANCEL);
}

void OrientDlg::AddOrientBtns()
{
	const int marginl = 10;
	const int marginr = 10;
	const int marginb = 10;
	const int spacing = 10;
	const int btnperrow = 4;
	const int btnwidth = 88;
	const int btnheight = 28;

	RECT pos;
	GetClientRect(&pos);
	const LONG clitowin = m_wHeight - pos.bottom;
	const LONG leftinit = pos.right - marginr - btnwidth * btnperrow - spacing * (btnperrow - 1);
	const LONG topinit = pos.bottom + spacing;
	const LONG rightedge = (pos.right + m_wWidth) / 2;
	pos.left = leftinit;
	pos.top = topinit;
	pos.right = pos.left + btnwidth;
	pos.bottom = pos.top + btnheight;

	CFont *pFont = GetFont();

	int i = 0;
	for(auto& ori : m_orients) {
		if (i >= MAX_ORIENTS)
			break;

		CButton *button = new CButton;
		if (button == NULL)
			break;
		if (i) {
			if ((i % btnperrow) == 0) {
				pos.left = leftinit;
				pos.right = leftinit + btnwidth;
				pos.top += btnheight + spacing;
				pos.bottom += btnheight + spacing;
			} else {
				pos.left += btnwidth + spacing;
				pos.right += btnwidth + spacing;
			}
		}
		if (!button->Create(ori.nameshort, WS_CHILD | WS_DISABLED | BS_PUSHBUTTON, pos, this, IDC_ORIENT0 + i)) {
			delete button;
			break;
		}
		button->SetFont(pFont);
		if (m_tooltip && (wcslen(ori.namefull)))
			m_toolTip.AddTool(button, ori.namefull);

		m_oribtns.push_back(button);
		i++;
	}

	if (i) {
		m_extBtn = new CButton;
		if (m_extBtn == NULL)
			return;
		m_addBtn = new CButton;
		if (m_addBtn == NULL)
			return;
		m_oriSplit = new CStatic;

		RECT btnrect;
		btnrect.left = marginl;
		btnrect.right = btnrect.left + btnheight;
		btnrect.top = topinit;
		btnrect.bottom = topinit + btnheight;

		if (!m_addBtn->Create(L"", WS_CHILD | WS_DISABLED | BS_CHECKBOX | BS_PUSHLIKE | BS_ICON, btnrect, this, IDC_ADDBTN)) {
			delete m_addBtn, m_addBtn = NULL;
		} else {
			m_addBtn->SetCheck(m_addtodoc);
			m_hDocArrOff = (HICON)LoadImage(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICONDOCARROFF), IMAGE_ICON, 24, 24, LR_DEFAULTCOLOR);
			m_hDocArrOn = (HICON)LoadImage(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICONDOCARRON), IMAGE_ICON, 24, 24, LR_DEFAULTCOLOR);
			if (m_hDocArrOff)
				m_addBtn->SetIcon(m_hDocArrOff);
			if (m_tooltip)
				m_toolTip.AddTool(m_addBtn, L"Добавлять в документ");
		}

		CButton* resbtn = (CButton*)GetDlgItem(IDC_RESET);
		if (resbtn == NULL)
			return;

		resbtn->GetWindowRect(&btnrect);
		btnrect.left -= spacing + btnheight;
		btnrect.right = btnrect.left + btnheight;
		ScreenToClient(&btnrect);
		if (!m_extBtn->Create(L"", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_ICON, btnrect, this, IDC_EXTBTN)) {
			delete m_extBtn, m_extBtn = NULL;
			return;
		}
		m_hExtDn = (HICON)LoadImage(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICONEXTDN), IMAGE_ICON, 24, 24, LR_DEFAULTCOLOR);
		m_hExtUp = (HICON)LoadImage(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDI_ICONEXTUP), IMAGE_ICON, 24, 24, LR_DEFAULTCOLOR);
		if (m_hExtDn)
			m_extBtn->SetIcon(m_hExtDn);
		if (m_tooltip)
			m_toolTip.AddTool(m_extBtn, L"Пользовательские ориентации");

		if (m_oriSplit) {
			btnrect.left = marginl;
			btnrect.top = m_wHeight - clitowin - 1;
			btnrect.right = rightedge - marginr;
			btnrect.bottom = btnrect.top;
			if (!m_oriSplit->Create(L"", WS_CHILD | SS_ETCHEDHORZ, btnrect, this, IDC_ORISPLIT))
				delete m_oriSplit, m_oriSplit = NULL;
		}

		m_wHeightExt = pos.bottom + marginb + clitowin;
	}
}

void OrientDlg::PlaceWindow(LONG height /* = LONG_MIN */)
{
	LONG wH = height > 0 ? height : m_wHeight;
	RECT rect;
	if ((m_autoplace != AUTOPLACE_DISABLED) || (m_wPosX == LONG_MIN) || (m_wPosY == LONG_MIN)) {
		int margin = 2;
		if ((m_autoplace < AUTOPLACE_TOPLEFT) || (m_autoplace > AUTOPLACE_BOTRIGHT))
			m_autoplace = AUTOPLACE_DEFAULT;

		::GetWindowRect(m_drawHwnd, &rect);
		switch(m_autoplace) {
		case AUTOPLACE_TOPLEFT:
			m_wPosX = rect.left + margin;
			m_wPosY = rect.top + margin;
			break;
		case AUTOPLACE_TOPRIGHT:
			m_wPosX = rect.right - m_wWidth - margin;
			m_wPosY = rect.top + margin;
			break;
		case AUTOPLACE_BOTLEFT:
			m_wPosX = rect.left + margin;
			m_wPosY = rect.bottom - wH - margin;
			break;
		case AUTOPLACE_BOTRIGHT:
			m_wPosX = rect.right - m_wWidth - margin;
			m_wPosY = rect.bottom - wH - margin;
			break;
		};	
	}
	SetWindowPos(NULL, m_wPosX, m_wPosY, m_wWidth, wH, SWP_NOZORDER);
}

void OrientDlg::UpdateControls(int sel /* =ALL */)
{
	if (sel & CTRLX) {
		m_angleX.SetData(m_curX);
		m_sliderX.SetPos(m_curX);
	}
	if (sel & CTRLY) {
		m_angleY.SetData(m_curY);
		m_sliderY.SetPos(m_curY);
	}
	if (sel & CTRLZ) {
		m_angleZ.SetData(m_curZ);
		m_sliderZ.SetPos(m_curZ);
	}
}

void OrientDlg::Orient()
{
	if ((m_vproj != NULL) && (m_place != NULL)) {
		C3DutilsApp::SetPlacementEulerDeg(m_place, m_curX, m_curY, m_curZ);
		m_vproj->SetCurrent();
	}
}

void OrientDlg::Preview(UINT delay /* = 0 */)
{
	if (m_preview) {
		if (delay)
			m_previewTimer = SetTimer(PREVIEW_TIMER, delay, NULL);
		else {
			KillTimer(m_previewTimer), m_previewTimer = 0;
			Orient();
		}
	} else if (m_previewTimer != 0)
		KillTimer(m_previewTimer), m_previewTimer = 0;
}

BOOL OrientDlg::AddDocProjection(const C3DutilsApp::orient_t& orient)
{
	LPCWSTR name = wcslen(orient.namefull) ? orient.namefull : name = orient.nameshort;

	m_vpcol->Refresh();
	IViewProjectionPtr vproj = m_vpcol->GetByName(const_cast<LPWSTR>(name), TRUE, FALSE);
	if ((vproj != NULL) && !m_vpcol->DetachByBody(vproj))
		return FALSE;

	vproj = m_vpcol->NewViewProjection();
	if (vproj == NULL)
		return FALSE;

	IPlacementPtr place = vproj->GetPlacement();
	if (place == NULL) {
		vproj->Release();
		return FALSE;
	}

	if (!C3DutilsApp::SetPlacementEulerDeg(place, orient.angleX, orient.angleY, orient.angleZ)) {
		vproj->Release();
		return FALSE;
	}

	vproj->SetName(const_cast<LPWSTR>(name));
	vproj->SetScale(1.0);
	return m_vpcol->Add(vproj);
}

/*void OrientDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}*/

BOOL OrientDlg::PreTranslateMessage(MSG* pMsg) 
  {
	if ( TranslateAccelerator( m_hWnd, m_hAccelTable, pMsg ) ) return TRUE;
	m_toolTip.RelayEvent(pMsg);

	return CDialog::PreTranslateMessage(pMsg);
}

BEGIN_MESSAGE_MAP(OrientDlg, CDialog)
	ON_WM_DESTROY()
	ON_WM_NCACTIVATE()
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDOK, &OrientDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &OrientDlg::OnBnClickedCancel)
	ON_WM_HSCROLL()
	ON_EN_KILLFOCUS(IDC_EDITX, &OrientDlg::OnEnKillfocusEditX)
	ON_EN_KILLFOCUS(IDC_EDITY, &OrientDlg::OnEnKillfocusEditY)
	ON_EN_KILLFOCUS(IDC_EDITZ, &OrientDlg::OnEnKillfocusEditZ)
	ON_NOTIFY(NE_UPDATE, IDC_EDITX, OnNotifyEditX)
	ON_NOTIFY(NE_UPDATE, IDC_EDITY, OnNotifyEditY)
	ON_NOTIFY(NE_UPDATE, IDC_EDITZ, OnNotifyEditZ)
	ON_STN_CLICKED(IDC_ICONX, &OrientDlg::OnStnClickedIconX)
	ON_STN_CLICKED(IDC_ICONY, &OrientDlg::OnStnClickedIconY)
	ON_STN_CLICKED(IDC_ICONZ, &OrientDlg::OnStnClickedIconZ)
	ON_BN_CLICKED(IDC_RESET, &OrientDlg::OnBnClickedReset)
	ON_BN_CLICKED(IDC_PREVIEW, &OrientDlg::OnBnClickedPreview)
	ON_BN_CLICKED(IDC_EXTBTN, &OrientDlg::OnBnClickedExtBtn)
	ON_BN_CLICKED(IDC_ADDBTN, &OrientDlg::OnBnClickedAddBtn)
	ON_CONTROL_RANGE(BN_CLICKED, IDC_ORIENT0, IDC_ORIENT0 + MAX_ORIENTS, OnBnClickedOrient)
	ON_WM_TIMER()
	ON_WM_MOVING()
	ON_WM_NCLBUTTONDBLCLK()
	ON_COMMAND(ID_ORIENTACCELERATOR_RETURN, &OrientDlg::OnBnClickedOk)
END_MESSAGE_MAP()

// OrientDlg message handlers
void OrientDlg::OnDestroy()
{
	CDialogEx::OnDestroy();

	for(auto& btn : m_oribtns) {
		if (btn) {
			btn->DestroyWindow();
			delete btn, btn = NULL;
		}
	}
	m_orients.empty();

	if (m_extBtn) {
		m_extBtn->DestroyWindow();
		delete m_extBtn, m_extBtn = NULL;
	}
	if (m_addBtn) {
		m_addBtn->DestroyWindow();
		delete m_addBtn, m_addBtn = NULL;
	}
	if (m_oriSplit) {
		m_oriSplit->DestroyWindow();
		delete m_oriSplit; m_oriSplit = NULL;
	}

	if (m_hExtUp)
		DestroyIcon(m_hExtUp), m_hExtUp = NULL;
	if (m_hExtDn)
		DestroyIcon(m_hExtDn), m_hExtDn = NULL;
	if (m_hDocArrOn)
		DestroyIcon(m_hDocArrOn), m_hDocArrOn = NULL;
	if (m_hDocArrOff)
		DestroyIcon(m_hDocArrOff), m_hDocArrOff = NULL;
}

BOOL OrientDlg::OnNcActivate(BOOL bActive)
{
	if (!bActive && m_active)
		Confirm();

	return TRUE;
}

void OrientDlg::OnClose()
{
	Cancel();
}

void OrientDlg::OnBnClickedOk()
{
	Confirm();
}

void OrientDlg::OnBnClickedCancel()
{
	Cancel();
}

void OrientDlg::OnHScroll(UINT nSBCode, UINT nPos, CScrollBar* pScrollBar)
{
	switch(pScrollBar->GetDlgCtrlID()) {
	case IDC_SLIDERX:
		m_angleX.SetData(m_curX = m_sliderX.GetPos());
		Preview(200);
		return;
	case IDC_SLIDERY:
		m_angleY.SetData(m_curY = m_sliderY.GetPos());
		Preview(200);
		return;
	case IDC_SLIDERZ:
		m_angleZ.SetData(m_curZ = m_sliderZ.GetPos());
		Preview(200);
		return;
	}
	CDialogEx::OnHScroll(nSBCode, nPos, pScrollBar);
}

void OrientDlg::OnEnKillfocusEditX()
{
	m_angleX.GetData(m_curX);
	m_sliderX.SetPos(m_curX);
	Preview();
}

void OrientDlg::OnEnKillfocusEditY()
{
	m_angleY.GetData(m_curY);
	m_sliderY.SetPos(m_curY);
	Preview();
}

void OrientDlg::OnEnKillfocusEditZ()
{
	m_angleZ.GetData(m_curZ);
	m_sliderZ.SetPos(m_curZ);
	Preview();
}

void OrientDlg::OnNotifyEditX(NMHDR* pNMHDR, LRESULT* pResult)
{
	m_angleX.GetData(m_curX);
	m_sliderX.SetPos(m_curX);
	m_previewBtn->SetCheck(BST_CHECKED), m_preview = TRUE;
	Preview();

	*pResult = 0;
}

void OrientDlg::OnNotifyEditY(NMHDR* pNMHDR, LRESULT* pResult)
{
	m_angleY.GetData(m_curY);
	m_sliderY.SetPos(m_curY);
	m_previewBtn->SetCheck(BST_CHECKED), m_preview = TRUE;
	Preview();

	*pResult = 0;
}

void OrientDlg::OnNotifyEditZ(NMHDR* pNMHDR, LRESULT* pResult)
{
	m_angleZ.GetData(m_curZ);
	m_sliderZ.SetPos(m_curZ);
	m_previewBtn->SetCheck(BST_CHECKED), m_preview = TRUE;
	Preview();

	*pResult = 0;
}

void OrientDlg::OnBnClickedReset()
{
	m_curX = m_initX;
	m_curY = m_initY;
	m_curZ = m_initZ;
	UpdateControls();

	Preview();
}

void OrientDlg::OnBnClickedPreview()
{
	m_preview = (m_previewBtn->GetCheck() == BST_CHECKED);
	Preview();
}

void OrientDlg::OnBnClickedExtBtn()
{
	m_extended = !m_extended;
	for(auto &btn : m_oribtns) {
		btn->ShowWindow(m_extended?SW_SHOW:SW_HIDE);
		btn->EnableWindow(m_extended);
	}
	if (m_oriSplit)
		m_oriSplit->ShowWindow(m_extended?SW_SHOW:SW_HIDE);
	m_addBtn->ShowWindow(m_extended?SW_SHOW:SW_HIDE);
	m_addBtn->EnableWindow(m_extended);
	m_extBtn->SetIcon(m_extended ? m_hExtUp : m_hExtDn);
	PlaceWindow(m_extended ? m_wHeightExt : m_wHeight);
}

void OrientDlg::OnBnClickedAddBtn()
{
	m_addtodoc = !m_addtodoc;
	if (m_addBtn) {
		m_addBtn->SetCheck(m_addtodoc);
		if (m_hDocArrOn && m_hDocArrOff)
			m_addBtn->SetIcon(m_addtodoc ? m_hDocArrOn : m_hDocArrOff);
	}
}

void OrientDlg::OnBnClickedOrient(UINT nID)
{
	size_t nBtn = nID - IDC_ORIENT0;
	if (nBtn >= m_orients.size())
		return;
	auto& orient = m_orients[nBtn];

	m_curX = orient.angleX;
	m_curY = orient.angleY;
	m_curZ = orient.angleZ;
	UpdateControls();

	if (m_addtodoc) {
		AddDocProjection(orient);
		Confirm();
	} else
		Preview();
}

void OrientDlg::OnStnClickedIconX()
{
	m_curX = 0.0;
	UpdateControls(CTRLX);
	Preview(400);
}

void OrientDlg::OnStnClickedIconY()
{
	m_curY = 0.0;
	UpdateControls(CTRLY);
	Preview(400);
}

void OrientDlg::OnStnClickedIconZ()
{
	m_curZ = 0.0;
	UpdateControls(CTRLZ);
	Preview(400);
}

void OrientDlg::OnTimer(UINT_PTR nIdEvent)
{
	if (nIdEvent == PREVIEW_TIMER) {
		m_previewTimer = 0;
		Orient();
	}
}

void OrientDlg::OnMoving(UINT nSide, LPRECT lpRect)
{
	CDialogEx::OnMoving(nSide, lpRect);

	m_wPosX = lpRect->left;
	m_wPosY = lpRect->top;
	m_autoplace = AUTOPLACE_DISABLED;
}

void OrientDlg::OnNcLButtonDblClk(UINT nFlags, CPoint point)
{
	LONG y = point.y - m_wPosY;
	if ((y > 0) && (y < m_wCaptHeight)) {
		m_autoplace = AUTOPLACE_DEFAULT;
		PlaceWindow(m_extended? m_wHeightExt : m_wHeight);
	}
}
