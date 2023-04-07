//
// NumericEdit.cpp
//
// Uptera - innovative reaction
// GHINI Mauro
//

#include "stdafx.h"

#ifndef __NUMERICEDIT_H__
#include "numericedit.h"
#endif //__NUMERICEDIT_H__

IMPLEMENT_DYNAMIC(CNumericEdit, CEdit);

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#ifndef _HUGE_ENUF
    #define _HUGE_ENUF  1e+300  // _HUGE_ENUF*_HUGE_ENUF must overflow
#endif

#ifndef INFINITY
#define INFINITY   ((float)(_HUGE_ENUF * _HUGE_ENUF))
#endif
#ifndef HUGE_VAL
#define HUGE_VAL   ((double)INFINITY)
#endif
#ifndef HUGE_VALF
#define HUGE_VALF  ((float)INFINITY)
#endif
#ifndef HUGE_VALL
#define HUGE_VALL  ((long double)INFINITY)
#endif
#ifndef NAN
#define NAN        ((float)(INFINITY * 0.0F))
#endif

//
// Gets Key state
//
SHORT GetAllKeyState()
{
	SHORT nKeyState;

	nKeyState = (SHORT)((GetKeyState(VK_SHIFT)<0   ? SHIFT_MASK : 0) |
						(GetKeyState(VK_CONTROL)<0 ? CTRL_MASK : 0) |
						(GetKeyState(VK_MENU)<0    ? ALT_MASK : 0));

	return nKeyState;
}


//
// NumericEdit implementation
//
CNumericEdit::CNumericEdit()
{
	// Set default
	m_strFormat = _T("%g");
	m_strPaste.Empty();
	m_nPastPoint = 0;
	m_dDefaultValue = 0.0;
	m_bNegativeInput = TRUE;
	m_bDecimalInput = TRUE;

	m_dMaxValue = HUGE_VAL;
	m_dMinValue = -HUGE_VAL;

	m_bKeyProcessed = FALSE;
	m_bForceProcessing = FALSE;

#ifdef _DEBUG
	m_bIsAttached = FALSE;
#endif
}

CNumericEdit::~CNumericEdit()
{
}

BEGIN_MESSAGE_MAP(CNumericEdit, CEdit)
	//{{AFX_MSG_MAP(CNumericEdit)
	ON_WM_KEYDOWN()
	ON_WM_CHAR()
	ON_MESSAGE(WM_PASTE, OnPaste)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

//
// RefreshDisplay
//
void CNumericEdit::RefreshDisplay(void)
{
	ASSERT_VALID(this);
	CString text;
	TCHAR* szStop;
	double dData;

	// format display string
	dData = 0.0;
	GetWindowText(text);
	dData = _tcstod(text, &szStop);
	text.Format(m_strFormat, dData);
	if (text.Find(_T("#")) != -1)
	{
		text.Format(m_strFormat, m_dDefaultValue);
		SetWindowText(text);
	}
}

//
// DisplayDefaultValue
//
void CNumericEdit::DisplayDefaultValue(void)
{
	ASSERT_VALID(this);

	CString text;
	text.Format(m_strFormat, m_dDefaultValue);
	SetWindowText(text);
}

//
// AttachEdit
//
// Subclasses an edit with ID nCtrlID and with parent pParentWnd
//
//   Uses CWnd::SubClassDlgItem which lets us insert
//   our message map into the MFC message routing system
//
BOOL CNumericEdit::AttachEdit(int nCtlID, CWnd* pParentWnd)
{
	ASSERT(pParentWnd != NULL);

	if (SubclassDlgItem(nCtlID, pParentWnd))
	{
#ifdef _DEBUG
		m_bIsAttached = TRUE;
#endif
		ASSERT_VALID(this);
		SetLimitText(15);
		return TRUE;
	}

	// SubClassDlgItem failed
	TRACE2("CNumericEdit, could not attach to %d %x\n", nCtlID, pParentWnd);
	ASSERT(FALSE);
	return FALSE;  
} 

//
// SetData
//
// Do not use SetWindowText directly, use this function instead
//
BOOL CNumericEdit::SetData(const double& dData)
{
	ASSERT_VALID(this);
	CString text;
	int bRet = TRUE;

	text.Format(m_strFormat, dData);
	if (text.Find(_T("#")) != -1)
	{
		text.Format(m_strFormat, m_dDefaultValue);
		bRet = FALSE;
	}
	SetWindowText(text);
	return bRet;
}

//
// SetData
//
BOOL CNumericEdit::SetData(const CString& strData)
{
	ASSERT_VALID(this);

	int nDataCount  = 0;                    //counter for the data string
	int nDataLen    = strData.GetLength();  //data length
	TCHAR chData;                           //data character buffer

	// if data contains valid char restore to default
	if (strData.FindOneOf(_T("0123456789eE.,-+")) == -1)
		return FALSE;

	// loop through the data string
	m_strPaste.Empty();
	m_nPastPoint = 0;
	CString strReplace;
	strReplace.Empty();
	for(nDataCount = 0; nDataCount < nDataLen; nDataCount++)
	{
		chData = strData.GetAt(nDataCount);
		// insert valid chars only
		if(IsPasteCharValid(chData))
			strReplace = strReplace + chData;
	}
	if (!strReplace.IsEmpty())
		CEdit::SetWindowText(strReplace);
	m_strPaste.Empty();

	return TRUE;
}

//
// GetData
//
// Get data from edit, call it if you want numeric value in current format
//
// Call GetWindowText if you want string value without any formatting
//
void CNumericEdit::GetData(double& dData)
{
	ASSERT_VALID(this);

	CString text;
	TCHAR* szStop;

	dData = 0.0;
	GetWindowText(text);
	// if text is empty set default value
	if (text.IsEmpty())
	{
		text.Format(m_strFormat, m_dDefaultValue);
	}
	dData = _tcstod(text, &szStop);
	text.Format(m_strFormat, dData);

	// check out of range
	if (dData <= -HUGE_VAL || dData >= HUGE_VAL)
	{
		if (text.Find(_T("-")) != -1)
		{
			dData = m_dMinValue;
		}
		else
		{
			dData = m_dMaxValue;
		}
	}

	// check Min/Max value
	if (dData <= m_dMinValue)
		dData = m_dMinValue;
	if (dData >= m_dMaxValue)
		dData = m_dMaxValue;

	// format display string
	text.Format(m_strFormat, dData);
	SetWindowText(text);
}

//
// PasteData
//
BOOL CNumericEdit::PasteData(const CString& strData)
{
	ASSERT_VALID(this);

	int nDataCount  = 0;                    //counter for the data string
	int nDataLen    = strData.GetLength();  //data length
	TCHAR chData;                           //data character buffer

	// if data contains valid char restore to default
	if (strData.FindOneOf(_T("0123456789eE.,-+")) == -1)
		return FALSE;

	// loop through the data string
	GetWindowText(m_strPaste);
	int nStart, nEnd;
	GetSel(nStart, nEnd);

	CString left;
	CString right;
	left = m_strPaste.Left(nStart);
	right = m_strPaste.Mid(nEnd);
	m_strPaste = left + right;
	m_nPastPoint = nStart;

	CString strReplace;
	strReplace.Empty();
	for(nDataCount = 0; nDataCount < nDataLen; nDataCount++)
	{		
		chData = strData.GetAt(nDataCount);
		// insert valid chars only
		if(IsPasteCharValid(chData))
			strReplace = strReplace + chData;
	}
	if (!strReplace.IsEmpty())
		CEdit::ReplaceSel(strReplace, TRUE);
	m_strPaste.Empty();

	return TRUE;
}

//
// IsCharValid
//
// Determines if the newly entered character, ch, is valid
//
BOOL CNumericEdit::IsCharValid(TCHAR& ch)
{
	ASSERT_VALID(this);

	BOOL bRet = FALSE;
	
	// check for char validity
	if (IsCharAlphaNumeric(ch) && !IsCharAlpha(ch))
	{
		bRet = TRUE;
	}
	else
	{
		switch ((unsigned char)ch)
		{
			case 8:     // backspace
			{
				return TRUE;
			}

			case '-':
			{
				// Check is made from IsNewStringValid
				bRet = TRUE;
				break;
			}

			case '+':
			{
				bRet = TRUE;
				break;
			}

			case '.':
			{
				if (GetDecimalInput())
					bRet = TRUE;
				break;
			}

			case 'e':
			case 'E':
			{
				bRet = TRUE;
				break;
			}
		}
	}

	if (bRet)
		bRet = IsNewStringValid(ch);

	return bRet;
}

//
// IsNewStringValid
//
// Checks if the newly entered character, ch, forms a valid string
//
BOOL CNumericEdit::IsNewStringValid(TCHAR& ch)
{
	CString s;
	GetWindowText(s);

	int nStart, nEnd;
	GetSel(nStart, nEnd);
	CString left;
	CString right;
	left = s.Left(nStart);
	right = s.Mid(nEnd);
	s = left + ch + right;
	s.MakeUpper();

	int nExpPos = s.Find(_T("E"));
	if (nExpPos != -1)
	{
		left = s.Left(nExpPos);
		right = s.Mid(nExpPos + 1);
	}
	else
	{
		left = s;
		right.Empty();
	}

	int i;
	int nNeg = 0;
	int nPos = 0;
	int nDec = 0;
	TCHAR c;

	// parse string left-part
	if (!left.IsEmpty())
	{
		c = left.GetAt(0);
		switch ((unsigned char)c)
		{
		case '-':
			nNeg = 100;
			break;
		case '+':
			nPos = 100;
			break;
		case ',':
		case '.':
			nDec++;
		}
		for (i=1; i < left.GetLength(); i++)
		{
			c = left.GetAt(i);
			switch ((unsigned char)c)
			{
			case '-':
				nNeg++;
				break;
			case '+':
				nPos++;
				break;
			case ',':
			case '.':
				nDec++;
			}
		}

		if (nNeg == 100 && GetNegativeInput() == FALSE)
			return FALSE;
		if ((nNeg > 0 && nNeg < 100) || nNeg > 100)
			return FALSE;
		if ((nPos > 0 && nPos < 100) || nPos > 100)
			return FALSE;
		if (nDec > 1)
			return FALSE;
	}

	if (right.IsEmpty())
		return TRUE;

	// parse string right-part
	int nExp = 0;
	nNeg = 0;
	nPos = 0;
	nDec = 0;
	c = right.GetAt(0);
	switch ((unsigned char)c)
	{
	case '-':
		nNeg = 100;
		break;
	case '+':
		nPos = 100;
		break;
	case ',':
	case '.':
		nDec++;
		break;
	case 'E':
		nExp++;
	}
	for (i=1; i < right.GetLength(); i++)
	{
		c = right.GetAt(i);
		switch ((unsigned char)c)
		{
		case '-':
			nNeg++;
			break;
		case '+':
			nPos++;
			break;
		case ',':
		case '.':
			nDec++;
			break;
		case 'E':
			nExp++;
		}
	}

	if (nNeg == 100 && GetDecimalInput() == FALSE)
		return FALSE;
	if ((nNeg > 0 && nNeg < 100) || nNeg > 100)
		return FALSE;
	if ((nPos > 0 && nPos < 100) || nPos > 100)
		return FALSE;
	if (nDec > 1 || nExp > 0)
		return FALSE;

	return TRUE;
}


//
// IsPasteCharValid
//
// Checks if the pasted character, ch, is valid
//
BOOL CNumericEdit::IsPasteCharValid(TCHAR& ch)
{
	ASSERT_VALID(this);

	BOOL bRet = FALSE;
	if (IsCharAlphaNumeric(ch) && !IsCharAlpha(ch))
	{
		bRet = TRUE;
	}
	else
	{
		switch ((unsigned char)ch)
		{
			case '-':
			{
				bRet = TRUE;
				break;
			}
			case '+':
			{
				bRet = TRUE;
				break;
			}
			case '.':
			{
				if (GetDecimalInput())
					bRet = TRUE;
				break;
			}
			case 'e':
			case 'E':
			{
				bRet = TRUE;
				break;
			}
		}
	}

	if (bRet)
		bRet = IsPasteStringValid(ch);

	return bRet;
}


//
// IsPasteStringValid
//
// Checks if the pasted character, ch, forms a valid string
//
BOOL CNumericEdit::IsPasteStringValid(TCHAR& ch)
{
	CString pastText;
	pastText = m_strPaste;
	if (pastText.GetLength() == 0)
	{
		m_strPaste = ch;
		m_nPastPoint++;
		return TRUE;
	}

	CString left;
	CString right;
	left = pastText.Left(m_nPastPoint);
	right = pastText.Mid(m_nPastPoint);
	pastText = left + ch + right;

	CString strOut = pastText;
	pastText.MakeUpper();

	int nExpPos = pastText.Find(_T("E"));
	if (nExpPos != -1)
	{
		left = pastText.Left(nExpPos);
		right = pastText.Mid(nExpPos + 1);
	}
	else
	{
		left = pastText;
		right.Empty();
	}

	int i;
	int nNeg = 0;
	int nPos = 0;
	int nDec = 0;
	TCHAR c;

	// parse string left-part
	if (!left.IsEmpty())
	{
		c = left.GetAt(0);
		switch ((unsigned char)c)
		{
		case '-':
			nNeg = 100;
			break;
		case '+':
			nPos = 100;
			break;
		case ',':
		case '.':
			nDec++;
		}
		for (i=1; i < left.GetLength(); i++)
		{
			c = left.GetAt(i);
			switch ((unsigned char)c)
			{
			case '-':
				nNeg++;
				break;
			case '+':
				nPos++;
				break;
			case ',':
			case '.':
				nDec++;
			}
		}

		if (nNeg == 100 && GetNegativeInput() == FALSE)
			return FALSE;
		if ((nNeg > 0 && nNeg < 100) || nNeg > 100)
			return FALSE;
		if ((nPos > 0 && nPos < 100) || nPos > 100)
			return FALSE;
		if (nDec > 1)
			return FALSE;
	}

	if (right.IsEmpty())
	{
		m_strPaste = strOut;
		m_nPastPoint++;
		return TRUE;
	}

	// parse string right-part
	int nExp = 0;
	nNeg = 0;
	nPos = 0;
	nDec = 0;
	c = right.GetAt(0);
	switch ((unsigned char)c)
	{
	case '-':
		nNeg = 100;
		break;
	case '+':
		nPos = 100;
		break;
	case ',':
	case '.':
		nDec++;
		break;
	case 'E':
		nExp++;
	}
	for (i=1; i < right.GetLength(); i++)
	{
		c = right.GetAt(i);
		switch ((unsigned char)c)
		{
		case '-':
			nNeg++;
			break;
		case '+':
			nPos++;
			break;
		case ',':
		case '.':
			nDec++;
			break;
		case 'E':
			nExp++;
		}
	}

	if (nNeg == 100 && GetDecimalInput() == FALSE)
		return FALSE;
	if ((nNeg > 0 && nNeg < 100) || nNeg > 100)
		return FALSE;
	if ((nPos > 0 && nPos < 100) || nPos > 100)
		return FALSE;
	if (nDec > 1 || nExp > 0)
		return FALSE;

	m_strPaste = strOut;
	m_nPastPoint++;
	return TRUE;
}

//
// OnKeydown
//
// Handles special keys (non WM_CHAR) such as delete, back, left, right tab, etc.
//
void CNumericEdit::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	ASSERT_VALID(this);

	m_bKeyProcessed = TRUE;
	m_bForceProcessing = FALSE;
	SHORT nKeyState = GetAllKeyState();

	// virtual keys
	switch(nChar)
	{
		case VK_SPACE:
			NMHDR nmh;
			nmh.code = NE_UPDATE;    // Message type defined by control.
			nmh.idFrom = GetDlgCtrlID();
			nmh.hwndFrom = m_hWnd;
			GetParent()->SendMessage(WM_NOTIFY, nmh.idFrom, (LPARAM)&nmh);
			SetSel(0, -1, TRUE);
			break;
		// standard GUI cut and paste command
		case 'V':
		{
			// CTRL+V ->Paste
			if (CTRL_MASK == nKeyState)
			{
				OnPaste(0, 0);
				break;
			}
		}
		case 'C':   // CTRL+C -> Copy
		case 'X':   // CTRL+X -> Cut
		case 'Z':   // CTRL+Z -> Undo
		{
			if (CTRL_MASK == nKeyState)
				m_bForceProcessing = TRUE;
		}

		default:
		{
			CEdit::OnKeyDown(nChar, nRepCnt, nFlags);
			m_bKeyProcessed = FALSE;
		}
	}
}

//
// OnChar
//
// Handles standard keys
//
void CNumericEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	ASSERT_VALID(this);

	if (IsReadOnly())
		return;

	if (m_bForceProcessing)
	{
		CEdit::OnChar(nChar, nRepCnt, nFlags);
		m_bForceProcessing = FALSE;
		return;
	}

	if(!m_bKeyProcessed)
	{
		if(IsCharValid((TCHAR&)nChar))
			CEdit::OnChar(nChar, nRepCnt, nFlags);
	}

	m_bKeyProcessed = FALSE;
}


//
// OnPaste
//
// Paste data from clipboard
//
afx_msg LRESULT CNumericEdit::OnPaste(WPARAM, LPARAM)
{
	if (IsReadOnly())
		return FALSE;

	// clipboard data available?
	if(!IsClipboardFormatAvailable(CF_TEXT))
		return 0;

	HMODULE hClipMem = NULL;
	LPSTR lpszData = NULL;	// pointer to data in clipboard

	//get the data from the clipboard
	if(OpenClipboard())
	{
		HGLOBAL handle;
#ifdef _UNICODE
		handle = GetClipboardData(CF_UNICODETEXT);
#else
		handle = GetClipboardData(CF_TEXT);

		if (handle == 0)
			handle = GetClipboardData(CF_OEMTEXT);
#endif

		if (handle)
		{
			LPTSTR psz = (LPTSTR)GlobalLock(handle);
			PasteData(CString(psz));
			GlobalUnlock(handle);
		}
		CloseClipboard();
	}

	return 1;
}

#ifdef _DEBUG

void CNumericEdit::Dump(CDumpContext& dc) const
{
	CWnd::Dump(dc);

	dc << "\nAttached = " << m_bIsAttached << "\n";
	dc << "\nPaste string = " << m_strPaste << "\n";
	dc << "\nFormat string = " << m_strFormat << "\n";
	dc << "\nDefault value = " << m_dDefaultValue << "\n";
	dc << "\nMin value = " << m_dMinValue << "\n";
	dc << "\nMax value = " << m_dMaxValue << "\n";
}

void CNumericEdit::AssertValid() const
{
	CWnd::AssertValid();
	if (m_bIsAttached == FALSE)
	{
		// Call CNumericEdit::AttachEdit BEFORE any member operations
		TRACE0("CNumericEdit: Performing operation on un-attached NumericEdit\n");
		ASSERT(FALSE);
	}
}

#endif //_DEBUG
