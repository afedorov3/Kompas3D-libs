////////////////////////////////////////////////////////////////////////////////
//
// GridViewDlg.cpp
// BatchExport library
// Alex Fedorov, 2018
//
////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "GridViewDlg.h"
#include "../common/document.h"
#include "../common/utils.h" // Str2Clipboard
#include <unordered_map>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#pragma warning(disable:4100)	// unreferenced formal parameter

#define BTN_MARGIN 5
#define COL_WIDTH_INCLUDE 25
#define COL_WIDTH_EMBODIMENT 90

GridViewDlg::GridViewDlg(LPCWSTR title, LPCWSTR OKBtnText, ksAPI7::IKompasDocument3DPtr doc3d,
		CBatchExportApp::VARIATIONS * variations, OPTIONS options /*= GVOPT_DEFAULT*/, GROUP_BY groupby /*= NO_GROUP*/, CWnd * pParent /*= NULL*/)
	: CDialog(GridViewDlg::IDD, pParent)
	, m_overwritefiles(BST_INDETERMINATE)
	, m_testmode(FALSE)
	, m_OKbtn(NULL)
	, m_Cancelbtn(NULL)
	, m_Savebtn(NULL)
	, m_Backbtn(NULL)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_title = title;
	m_doc3d = doc3d;
	m_okbtntext = OKBtnText;
	m_variations = variations;
	m_options = options;
	m_groupby = groupby;
	m_hAccelTable = LoadAccelerators(AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_ACCELERATORS));
}

void GridViewDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_GRID, m_ListCtrl);
}

BOOL GridViewDlg::PreTranslateMessage(MSG* pMsg) 
  {
	if ( TranslateAccelerator( m_hWnd, m_hAccelTable, pMsg ) ) return TRUE;
	m_ToolTip.RelayEvent(pMsg);

	return CDialog::PreTranslateMessage(pMsg);
}

BEGIN_MESSAGE_MAP(GridViewDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_GETMINMAXINFO()
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_NOTIFY(LVN_COLUMNCLICK, IDC_GRID, OnHeaderClick)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_GRID, OnLvnItemchangedList)
	ON_MESSAGE(WM_EMBODIMENTCHANGE, OnEmbodimentChange)
	ON_MESSAGE(WM_DIRCHANGE, OnDirChange)
	ON_MESSAGE(WM_NAMECHANGE, OnNameChange)
	ON_MESSAGE(WM_COMMENTADDCHANGE, OnCommentaddChange)
	ON_MESSAGE(WM_VARIABLESCHANGE, OnVariablesChange)
	ON_MESSAGE(WM_OPTIONSCHANGE, OnOptionsChange)
	ON_BN_CLICKED(IDC_SAVE, OnBnClickedSave)
	ON_BN_CLICKED(IDC_BACK, OnBnClickedBack)
	ON_BN_CLICKED(IDC_OVERWRITE, OnBnClickedOverwrite)
	ON_BN_CLICKED(IDC_TESTMODE, OnBnClickedTestMode)
	ON_COMMAND(ID_ACCELERATOR_CTRLINS, ClipboardCopy)
END_MESSAGE_MAP()

// GridViewDlg message handlers

BOOL GridViewDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	m_OKbtn = (CButton*)GetDlgItem(IDOK);
	m_Cancelbtn = (CButton*)GetDlgItem(IDCANCEL);
	m_Savebtn = (CButton*)GetDlgItem(IDC_SAVE);
	m_Backbtn = (CButton*)GetDlgItem(IDC_BACK);
	m_Overwritechk = (CButton*)GetDlgItem(IDC_OVERWRITE);
	m_TestModechk = (CButton*)GetDlgItem(IDC_TESTMODE);
	if (HasOpts(GVOPT_SHOWFILEOPTS)) {
		m_Overwritechk->SetCheck(m_overwritefiles);
		m_Overwritechk->EnableWindow(!m_testmode);
		m_Overwritechk->ShowWindow(SW_SHOW);
		m_TestModechk->SetCheck(m_testmode);
		m_TestModechk->ShowWindow(SW_SHOW);
	}
	if (HasOpts(GVOPT_SHOWCANCEL))
		m_Cancelbtn->ShowWindow(SW_SHOW);
	if (HasOpts(GVOPT_CANBACK) && (HasOpts(GVOPT_READONLY) || NoOpts(GVOPT_ALLOWSAVE)))
		m_Backbtn->ShowWindow(SW_SHOW);

	// Give better margin to editors
	m_ListCtrl.SetCellMargin(1.4);

	// enable extended styles
	m_ListCtrl.SetExtendedStyle(m_ListCtrl.GetExtendedStyle() | LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT | LVS_EX_HIDELABELS );
	CRect listrc;
	m_ListCtrl.GetClientRect(&listrc);
	for (INT col = 0; col < CBatchExportApp::FIELD_COUNT + 1; col++) { // + Checkbox column
		CGridColumnTrait *pTrait = NULL;
		INT width = 150;
		LPCWSTR name = L"";
		if (col > 0)
			name = CBatchExportApp::fieldnames[col - 1];
		if (col == (CBatchExportApp::VP_EMBODIMENT + 1)) {
			CGridColumnTraitCombo* pComboTrait = new CGridColumnTraitCombo;
			pComboTrait->SetSingleClickEdit(true);
			INT emcount = Document::EmbodimentCount(m_doc3d);
			CString str;
			for (INT i = 0; i < emcount; i++) {
				_itot_s( i, str.GetBufferSetLength( 40 ), 40, 10 );
				str.ReleaseBuffer();
				pComboTrait->AddItem((DWORD_PTR)i, str);
				pComboTrait->SetStyle(WS_VSCROLL | CBS_DROPDOWNLIST );
			}
			pTrait = pComboTrait;
		} else
			pTrait = new CGridColumnTraitEdit;
		CGridColumnTrait::ColumnState *state = &pTrait->GetColumnState();
		state->m_Editable = (bool)NoOpts(GVOPT_READONLY);
		if (col == 0) {
			if (NoOpts(GVOPT_SHOWINCLUDE))
				state->m_AlwaysHidden = true;
			state->m_Resizable = false;
			width = COL_WIDTH_INCLUDE;
		} else if (col == (CBatchExportApp::VP_EMBODIMENT + 1))
			width = COL_WIDTH_EMBODIMENT;
		else if ((col == (CBatchExportApp::VP_VARIABLES + 1)) && NoOpts(GVOPT_SHOWVARIABLES))
			state->m_AlwaysHidden = true;
		else if ((col == (CBatchExportApp::VP_STATUS + 1)) && NoOpts(GVOPT_SHOWSTATUS))
			state->m_AlwaysHidden = true;
		else
			width = 100;
		state->m_Sortable = false;
		state->m_MinWidth = 30;
		m_ListCtrl.InsertColumnTrait(col, name, LVCFMT_LEFT, width, col - 1, pTrait);
	}
	CHeaderCtrl *hdrCtrl = m_ListCtrl.GetHeaderCtrl();
	if (hdrCtrl == NULL)
		return FALSE;
	hdrCtrl->ModifyStyle(HDS_DRAGDROP, 0);

	switch(m_groupby) {
	case GROUP_BY_ATTR:
	case GROUP_BY_ATTRROW:
		m_ListCtrl.EnableGroupView( TRUE );
		if (!m_ListCtrl.IsGroupViewEnabled())
			m_groupby = NO_GROUP;
		break;
	case NO_GROUP:
		break;
	default:
		m_groupby = NO_GROUP;
		break;
	}

	// populate list data
	size_t varcount = m_variations->size();
	CString strCellText;
	INT nItem = 0;
	// groupping stuff
	std::unordered_map<reference, LONG> aids;
	std::unordered_map<UINT64, INT> gids;
	LONG attrindex = 0;
	INT groupindex = 0;
	for(size_t vari = 0; vari < varcount ; vari++)
	{
		CBatchExportApp::VARIATION *varptr = &m_variations->at(vari);
		BOOL include = FALSE;
		if ((varptr->options & CBatchExportApp::VAROPT_INCLUDE) == CBatchExportApp::VAROPT_INCLUDE) {
			include = TRUE;
		} else if (HasOpts(GVOPT_SKIPEXCL))
			continue;
		if ((nItem = m_ListCtrl.InsertItem(++nItem, L"")) < 0) // will be updated later
			break;
		m_ListCtrl.SetItemData(nItem, (DWORD_PTR)varptr);
		// embodiment
		_itot_s(varptr->embodiment , strCellText.GetBufferSetLength( 40 ), 40, 10 );
		strCellText.ReleaseBuffer();
		m_ListCtrl.SetItemText(nItem, CBatchExportApp::VP_EMBODIMENT + 1, strCellText);
		// dir
		m_ListCtrl.SetItemText(nItem, CBatchExportApp::VP_DIR + 1, varptr->dir);
		// name
		m_ListCtrl.SetItemText(nItem, CBatchExportApp::VP_NAME + 1, varptr->name);
		// commentadd
		m_ListCtrl.SetItemText(nItem, CBatchExportApp::VP_COMMENTADD + 1, varptr->commentadd);
		// variables
		m_ListCtrl.SetItemText(nItem, CBatchExportApp::VP_VARIABLES + 1, varptr->varstr);
		// options & include
		ItemInclude(nItem, include);
		// status
		CString auxstr;
		CBatchExportApp::StrStatus(auxstr, varptr->status);
		m_ListCtrl.SetItemText(nItem, CBatchExportApp::VP_STATUS + 1, (LPCWSTR)auxstr);

		if (m_groupby != NO_GROUP) {
			UINT64 groupselect = 0;
			switch(m_groupby) {
			case GROUP_BY_ATTRROW:
				groupselect |= varptr->AttrRow;
			case GROUP_BY_ATTR:
				groupselect |= (UINT64)varptr->AttrPtr << 32;
			}
			INT &rgid = gids[groupselect];
			if (rgid == 0) {
				rgid = ++groupindex;
				LVGROUP lg = {0};
				lg.cbSize = sizeof(lg);
				lg.iGroupId = rgid;
				lg.state = LVGS_NORMAL;
				lg.mask = LVGF_GROUPID | LVGF_HEADER | LVGF_STATE | LVGF_ALIGN;
				lg.uAlign = LVGA_HEADER_LEFT;
 
				LONG &raid = aids[varptr->AttrPtr];
				if (raid == 0)
					raid = ++attrindex;
				switch(m_groupby) {
				case GROUP_BY_ATTRROW:
					auxstr.Format(L"Атрибут %ld, шаблон %ld", raid, varptr->AttrRow + 1);
					break;
				case GROUP_BY_ATTR:
					auxstr.Format(L"Атрибут %ld", raid);
					break;
				}
				lg.pszHeader = auxstr.GetBuffer();
				lg.cchHeader = auxstr.GetLength();
				if (m_ListCtrl.InsertGroup(rgid, &lg) == -1) {	// disable group view on any group manipulation fail
					m_groupby = NO_GROUP;
					m_ListCtrl.EnableGroupView( FALSE );
					continue;
				}
			}
			LVITEM lvItem = {0};
			lvItem.mask = LVIF_GROUPID;
			lvItem.iItem = nItem;
			lvItem.iSubItem = 0;
			lvItem.iGroupId = rgid;
			m_ListCtrl.SetItem( &lvItem );
		}
	}

	// titles and captions
	CString title = m_title;
	title.AppendFormat(L" (%d)", nItem + 1);
	SetWindowText((LPCWSTR)title);
	m_OKbtn->SetWindowText(HasOpts(GVOPT_SHOWFILEOPTS) && m_testmode?L"Тест":(LPCWSTR)m_okbtntext);

	// controls size and position
	CRect rc;
	GetClientRect(&rc);
	MoveControls(&rc);
	ForbidSave();

	// ToolTip control
	if (m_ToolTip.Create(this)) {
		m_ToolTip.AddTool(m_TestModechk, L"Выполнять всё, кроме файловых операций");
		m_ToolTip.Activate(TRUE);

		m_ListCtrl.AddHeaderToolTip(0, L"Включить/исключить все, с SHIFT - инвертировать выбор");
	}

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void GridViewDlg::OnGetMinMaxInfo(MINMAXINFO FAR* lpMMI)
{
  // set the minimum tracking width
  // and the minimum tracking height of the window
  lpMMI->ptMinTrackSize.x = 400;
  lpMMI->ptMinTrackSize.y = 200;
}

void GridViewDlg::OnSize( UINT nType, int cx, int cy )
{
	CDialog::OnSize(nType, cx, cy);

	CRect rc;
    GetClientRect(&rc);
	MoveControls(&rc);
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.
void GridViewDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		INT cxIcon = GetSystemMetrics(SM_CXICON);
		INT cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		INT x = (rect.Width() - cxIcon + 1) / 2;
		INT y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR GridViewDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void GridViewDlg::OnBnClickedSave()
{
	if (NoOpts(GVOPT_ALLOWSAVE) || HasOpts(GVOPT_READONLY))
		return;
	ForbidSave();
	DisableOptions(GVOPT_ALLOWSAVE);
	theApp.VariationsSave( m_doc3d, m_variations );
	EnableOptions(GVOPT_ALLOWSAVE);
}

void GridViewDlg::OnBnClickedBack()
{
	EndDialog(IDC_BACK);
}

void GridViewDlg::OnBnClickedOverwrite()
{
	m_overwritefiles = m_Overwritechk->GetCheck();
}

void GridViewDlg::OnBnClickedTestMode()
{
	m_testmode = m_TestModechk->GetCheck();
	m_Overwritechk->EnableWindow(!m_testmode);
	m_OKbtn->SetWindowText(m_testmode?L"Тест":(LPCWSTR)m_okbtntext);
}

void GridViewDlg::OnHeaderClick(NMHDR* pNMHDR, LRESULT* pResult)
{
	if (NoOpts(GVOPT_SHOWINCLUDE))
		return;

	LPNMLISTVIEW pLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);

	INT nCol = pLV->iSubItem;
	if (nCol == 0) {
		int varsz = (int)m_variations->size();
		int nRow = 0;
		if (varsz < 1) {
			*pResult = 0;
			return;
		}
		ULONG included = CBatchExportApp::VAROPT_INCLUDE;
		CBatchExportApp::VARIATION *varptr = NULL;
		if (m_ListCtrl.GetSelectedCount() > 1) {
			// process only selected items
			if (GetKeyState(VK_SHIFT) < 0) { // SHIFT is down, do inversion
				POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
				while (pos != NULL) {
					INT nRow = m_ListCtrl.GetNextSelectedItem(pos);
					varptr = &m_variations->at(nRow);
					if (varptr->status != LIBSTATUS_SUCCESS)
						continue;
					ULONG iteminclude = varptr->options & CBatchExportApp::VAROPT_INCLUDE;
					m_ListCtrl.SetCheck(nRow, iteminclude?FALSE:TRUE);
					included &= iteminclude;
				}
			} else {
				POSITION pos = m_ListCtrl.GetFirstSelectedItemPosition();
				while (pos != NULL) {
					INT nRow = m_ListCtrl.GetNextSelectedItem(pos);
					varptr = &m_variations->at(nRow);
					if (varptr->status != LIBSTATUS_SUCCESS)
						continue;
					included &= varptr->options & CBatchExportApp::VAROPT_INCLUDE;
				}
				pos = m_ListCtrl.GetFirstSelectedItemPosition();
				while (pos != NULL) {
					INT nRow = m_ListCtrl.GetNextSelectedItem(pos);
					varptr = &m_variations->at(nRow);
					if (varptr->status != LIBSTATUS_SUCCESS)
						continue;
					m_ListCtrl.SetCheck(nRow, included?FALSE:TRUE);
				}
			}
		} else {
			// process all items
			if (GetKeyState(VK_SHIFT) < 0) { // SHIFT is down, do inversion
				for(nRow = 0; nRow < varsz; nRow++) {
					varptr = &m_variations->at(nRow);
					if (varptr->status != LIBSTATUS_SUCCESS)
						continue;
					ULONG iteminclude = varptr->options & CBatchExportApp::VAROPT_INCLUDE;
					m_ListCtrl.SetCheck(nRow, iteminclude?FALSE:TRUE);
					included &= iteminclude;
				}
			} else {
				// scan through all elements to deternime operation
				for(nRow = 0; nRow < varsz; nRow++) {
					varptr = &m_variations->at(nRow);
					if (varptr->status != LIBSTATUS_SUCCESS)
						continue;
					included &= varptr->options & CBatchExportApp::VAROPT_INCLUDE;
				}
				for(nRow = 0; nRow < varsz; nRow++) {
					varptr = &m_variations->at(nRow);
					if (varptr->status != LIBSTATUS_SUCCESS)
						continue;
					m_ListCtrl.SetCheck(nRow, included?FALSE:TRUE);
				}
			}
		}
		EnableOK(!included);
		AllowSave();
	}
}

void GridViewDlg::OnLvnItemchangedList(NMHDR * pNMHDR, LRESULT * pResult)
{
	if (NoOpts(GVOPT_SHOWINCLUDE))
		return;

	LPNMLISTVIEW pLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);

	if(pLV->uChanged & LVIF_STATE) { // item state has been changed
		int nRow = pLV->iItem;
		if ((pLV->uOldState == 0) || (nRow < 0) || !(nRow < (int)m_variations->size())) {
			*pResult = 0;
			return;
		}
		CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);
		BOOL include = FALSE;
		switch(pLV->uNewState & 0x3000) {
		case 0x2000: // new state: checked
			if (ItemInclude(nRow, TRUE))
				AllowSave();
			break;
		case 0x1000: // new state: unchecked
			ItemInclude(nRow, FALSE);
			AllowSave();
			break;
		}
	}

	int varsz = (int)m_variations->size();
	int nRow;
	ULONG included = 0;
	CBatchExportApp::VARIATION *varptr = NULL;
	for(nRow = 0; nRow < varsz; nRow++) {
		varptr = &m_variations->at(nRow);
		if (varptr->status != LIBSTATUS_SUCCESS)
			continue;
		included |= varptr->options & CBatchExportApp::VAROPT_INCLUDE;
	}
	EnableOK(included);

	*pResult = 0;
}

LRESULT GridViewDlg::OnEmbodimentChange(WPARAM wParam, LPARAM lParam)
{
	INT nRow = (INT)wParam;
	LPCWSTR pszText = (LPCWSTR)lParam;
	if (nRow > (INT)m_variations->size())
		return (LRESULT)FALSE;

	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);
	LPWSTR endptr;
	LONG val = wcstol(pszText, &endptr, 0);
	if ((endptr == NULL) || (*endptr != '\0') ||
		(val < 0) || (val == LONG_MAX))
		return (LRESULT)FALSE;
	varptr->embodiment = val;

	varptr->status &= ~LIBSTATUS_EMBODIMENT_MASK; // reset embodiment status
	if (varptr->embodiment >= Document::EmbodimentCount(m_doc3d))
		varptr->status |= LIBSTATUS_BAD_EMBODIMENT;
	else if (HasOpts(GVOPT_PRESERVETEMPL)) { // only recheck other fields when templates are preserved
		// Dir
		varptr->status &= ~LIBSTATUS_DIR_MASK;
		CString oldval(varptr->dir);
		varptr->status |= theApp.ComposeDir(m_doc3d, oldval, varptr);
		varptr->dir.SetString(oldval); // restore as template
		// Name
		varptr->status &= ~LIBSTATUS_NAME_MASK;
		oldval = varptr->name;
		varptr->status |= theApp.ComposeString(m_doc3d, oldval, varptr, CBatchExportApp::VP_NAME);
		varptr->name.SetString(oldval);
		// Comment
		varptr->status &= ~LIBSTATUS_COMMENTADD_MASK;
		oldval = varptr->commentadd;
		varptr->status |= theApp.ComposeString(m_doc3d, oldval, varptr, CBatchExportApp::VP_COMMENTADD);
		varptr->commentadd.SetString(oldval);
		// Vars
		varptr->status &= ~LIBSTATUS_VARS_MASK;
		varptr->status |= theApp.ParseVariables(m_doc3d, (LPCWSTR)varptr->varstr, varptr);
	}

	m_ListCtrl.SetCheck(nRow, varptr->status == LIBSTATUS_SUCCESS);

	AllowSave();
	return (LRESULT)TRUE;
}

LRESULT GridViewDlg::OnDirChange(WPARAM wParam, LPARAM lParam)
{
	INT nRow = (INT)wParam;
	LPCWSTR pszText = (LPCWSTR)lParam;
	if (nRow > (INT)m_variations->size())
		return (LRESULT)FALSE;

	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);

	varptr->status &= ~LIBSTATUS_DIR_MASK;
	varptr->status |= theApp.ComposeDir(m_doc3d, pszText, varptr);
	if (HasOpts(GVOPT_PRESERVETEMPL))
		varptr->dir.SetString(pszText); // restore as template
	m_ListCtrl.SetCheck(nRow, varptr->status == LIBSTATUS_SUCCESS);

	AllowSave();

	return (LRESULT)TRUE;
}

LRESULT GridViewDlg::OnNameChange(WPARAM wParam, LPARAM lParam)
{
	INT nRow = (INT)wParam;
	LPCWSTR pszText = (LPCWSTR)lParam;
	if (nRow > (INT)m_variations->size())
		return (LRESULT)FALSE;

	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);

	varptr->status &= ~LIBSTATUS_NAME_MASK;
	varptr->status |= theApp.ComposeString(m_doc3d, pszText, varptr, CBatchExportApp::VP_NAME);
	if (HasOpts(GVOPT_PRESERVETEMPL))
		varptr->name.SetString(pszText);
	m_ListCtrl.SetCheck(nRow, varptr->status == LIBSTATUS_SUCCESS);

	AllowSave();
	return (LRESULT)TRUE;
}

LRESULT GridViewDlg::OnCommentaddChange(WPARAM wParam, LPARAM lParam)
{
	INT nRow = (INT)wParam;
	LPCWSTR pszText = (LPCWSTR)lParam;
	if (nRow > (INT)m_variations->size())
		return (LRESULT)FALSE;

	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);

	varptr->status &= ~LIBSTATUS_COMMENTADD_MASK;
	varptr->status |= theApp.ComposeString(m_doc3d, pszText, varptr, CBatchExportApp::VP_COMMENTADD);
	if (HasOpts(GVOPT_PRESERVETEMPL))
		varptr->commentadd.SetString(pszText);
	m_ListCtrl.SetCheck(nRow, varptr->status == LIBSTATUS_SUCCESS);

	AllowSave();
	return (LRESULT)TRUE;
}

LRESULT GridViewDlg::OnVariablesChange(WPARAM wParam, LPARAM lParam)
{
	INT nRow = (INT)wParam;
	LPCWSTR pszText = (LPCWSTR)lParam;
	if (nRow > (INT)m_variations->size())
		return (LRESULT)FALSE;

	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);

	varptr->varstr.SetString(pszText);
	varptr->status &= ~LIBSTATUS_VARS_MASK;
	varptr->status |= theApp.ParseVariables(m_doc3d, (LPCWSTR)varptr->varstr, varptr);
	m_ListCtrl.SetCheck(nRow, varptr->status == LIBSTATUS_SUCCESS);

	AllowSave();
	return (LRESULT)TRUE;
}

LRESULT GridViewDlg::OnOptionsChange(WPARAM wParam, LPARAM lParam)
{
	INT nRow = (INT)wParam;
	LPCWSTR pszText = (LPCWSTR)lParam;
	if (nRow > (INT)m_variations->size())
		return (LRESULT)FALSE;

	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);

	BOOL ret = TRUE;
	if (NoOpts(GVOPT_NOCOMPINC)) {
		varptr->status &= ~LIBSTATUS_OPTS_MASK;
		varptr->status |= CBatchExportApp::ParseOptions(pszText, &varptr->options);
		BOOL state = (varptr->options & CBatchExportApp::VAROPT_INCLUDE) == CBatchExportApp::VAROPT_INCLUDE;
		ItemInclude(nRow, (varptr->options & CBatchExportApp::VAROPT_INCLUDE) == CBatchExportApp::VAROPT_INCLUDE);

		ret = FALSE; // we're updating the item ourselves
	} else
		varptr->optstr.SetString(pszText);

	AllowSave();

	return (LRESULT)ret;
}

void GridViewDlg::ClipboardCopy()
{
	CString clip;
	INT rowcount = m_ListCtrl.GetItemCount();
	INT colcount = m_ListCtrl.GetColumnCount();
	// header
	for(INT nCol = 0; nCol < colcount; nCol++) {
		if (!m_ListCtrl.IsColumnVisible(nCol))
			continue;
		if (nCol > 0)
			clip.AppendChar('\t');
		if (nCol == 0)
			clip.Append(L"Включено");
		else
			clip.Append(m_ListCtrl.GetColumnHeading(nCol));
	}
	clip.AppendChar('\n');
	// data
	for(INT nRow = 0; nRow < rowcount; nRow++) {
		for(INT nCol = 0; nCol < colcount; nCol++) {
			if (!m_ListCtrl.IsColumnVisible(nCol))
				continue;
			if (nCol > 0)
				clip.AppendChar('\t');
			if (nCol == 0) {
				if (m_ListCtrl.GetCheck(nRow))
					clip.Append(L"TRUE");
				else
					clip.Append(L"FALSE");
			} else
				clip.Append(m_ListCtrl.GetItemText(nRow, nCol));
			
		}
		clip.AppendChar('\n');
	}

	Utils::Str2Clipboard((LPCWSTR)clip);
}

BOOL GridViewDlg::ItemInclude(UINT nRow, BOOL include)
{
	if (nRow > (INT)m_variations->size())
		return (LRESULT)FALSE;
	CBatchExportApp::VARIATION *varptr = (CBatchExportApp::VARIATION *)m_ListCtrl.GetItemData(nRow);
	if ((varptr->status != LIBSTATUS_SUCCESS) && NoOpts(GVOPT_ERRIGNORE))
		include = FALSE; // force to excluded if in error state
	if (include) {
		m_ListCtrl.SetItemText(nRow, 0, L"Включено");
		varptr->options |= CBatchExportApp::VAROPT_INCLUDE;
	} else {
		m_ListCtrl.SetItemText(nRow, 0, L"Исключено");
		varptr->options &= ~CBatchExportApp::VAROPT_INCLUDE;
	}
	if (NoOpts(GVOPT_NOCOMPINC))
		CBatchExportApp::ComposeOptions(varptr->options, &varptr->optstr);
	m_ListCtrl.SetItemText(nRow, CBatchExportApp::VP_OPTIONS + 1, varptr->optstr);
	m_ListCtrl.SetCheck(nRow, include);

	return include;
}

void GridViewDlg::MoveControls(CRect * rc)
{
	CRect OKrc, Cancelrc, Saverc, Ovrrc, TMrc;

	if (!m_OKbtn)
		return;

	m_OKbtn->GetClientRect(&OKrc);
	m_Cancelbtn->GetClientRect(&Cancelrc);
	m_Savebtn->GetClientRect(&Saverc);
	m_Overwritechk->GetClientRect(&Ovrrc);
	m_TestModechk->GetClientRect(&TMrc);

	// controls position and size
	m_ListCtrl.SetWindowPos(NULL, 0, 0, rc->Width(), rc->Height() - OKrc.Height() - BTN_MARGIN*2, SWP_NOACTIVATE|SWP_NOZORDER);
	int ccount = m_ListCtrl.GetColumnCount();
	if (ccount > 0) {
		CRect listrc;
		m_ListCtrl.GetClientRect(&listrc);
		INT wcount = CBatchExportApp::FIELD_COUNT - 1 - 2; // minus checkboxes, embodiment
		INT varwidth = listrc.Width() - COL_WIDTH_EMBODIMENT;
		if (HasOpts(GVOPT_SHOWVARIABLES))
			wcount++;
		if (HasOpts(GVOPT_SHOWSTATUS))
			wcount++;
		if (HasOpts(GVOPT_SHOWINCLUDE))
			varwidth -= COL_WIDTH_INCLUDE;
		LONG columnwidth = varwidth / wcount;
		for (int i = 2; i < ccount; i++)
			m_ListCtrl.SetColumnWidth(i, columnwidth);
	}
	// options
	LONG cx = BTN_MARGIN;
	LONG cy = rc->Height() - (OKrc.Height() + Ovrrc.Height()) / 2 - BTN_MARGIN;
	m_Overwritechk->SetWindowPos(NULL, cx, cy, 0, 0, SWP_NOACTIVATE|SWP_NOZORDER|SWP_NOSIZE);
	cx += Ovrrc.Width() + BTN_MARGIN;
	m_TestModechk->SetWindowPos(NULL, cx, cy, 0, 0, SWP_NOACTIVATE|SWP_NOZORDER|SWP_NOSIZE);
	// buttons
	cx = rc->Width() - OKrc.Width() - BTN_MARGIN;
	cy = rc->Height() - OKrc.Height() - BTN_MARGIN;
	m_OKbtn->SetWindowPos(NULL, cx, cy, 0, 0, SWP_NOACTIVATE|SWP_NOZORDER|SWP_NOSIZE);
	cx -= Cancelrc.Width() + BTN_MARGIN;
	m_Cancelbtn->SetWindowPos(NULL, cx, cy, 0, 0, SWP_NOACTIVATE|SWP_NOZORDER|SWP_NOSIZE);
	cx -= Saverc.Width() + BTN_MARGIN;
	m_Savebtn->SetWindowPos(NULL, cx, cy, 0, 0, SWP_NOACTIVATE|SWP_NOZORDER|SWP_NOSIZE);
	m_Backbtn->SetWindowPos(NULL, cx, cy, 0, 0, SWP_NOACTIVATE|SWP_NOZORDER|SWP_NOSIZE);
	m_OKbtn->RedrawWindow();
	if (HasOpts(GVOPT_SHOWCANCEL))
		m_Cancelbtn->RedrawWindow();
	if (HasOpts(GVOPT_ALLOWSAVE) && NoOpts(GVOPT_READONLY))
		m_Savebtn->RedrawWindow();
	else if (HasOpts(GVOPT_CANBACK))
		m_Backbtn->RedrawWindow();
	if (HasOpts(GVOPT_SHOWFILEOPTS)) {
		m_Overwritechk->RedrawWindow();
		m_TestModechk->RedrawWindow();
	}
}

void GridViewDlg::AllowSave()
{
	if (HasOpts(GVOPT_ALLOWSAVE) && NoOpts(GVOPT_READONLY))
		m_Savebtn->ShowWindow(SW_SHOW);
}

void GridViewDlg::ForbidSave()
{
	m_Savebtn->ShowWindow(SW_HIDE);
}

void GridViewDlg::EnableOK(UINT enable)
{
	m_OKbtn->EnableWindow(enable);
}
