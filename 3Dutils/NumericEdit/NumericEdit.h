//
// NumericEdit.h
//
// Uptera - innovative reaction
// GHINI Mauro
//

#ifndef __NUMERICEDIT_H__
#define __NUMERICEDIT_H__

SHORT GetAllKeyState(void);

#define SHIFT_MASK  0x01
#define CTRL_MASK   0x02
#define ALT_MASK    0x04
#define NE_UPDATE	0x00

//
// NumericEdit declaration
//
class CNumericEdit : public CEdit
{
	DECLARE_DYNAMIC(CNumericEdit);

public:
	CNumericEdit();
	virtual ~CNumericEdit();

public:
	// Attach to edit control
	BOOL AttachEdit(int nCtlID, CWnd* pParentWnd);

	// Sets and Gets data
	virtual BOOL SetData(const double& dData);
	virtual BOOL SetData(const CString& strData);
	virtual void GetData(double& dData);

public:
	// Default value
	virtual void SetDefaultValue(double d)
	{
		m_dDefaultValue = d;
	}
	virtual double GetDefaultValue() const
	{
		return m_dDefaultValue;
	}

	// Format string
	virtual void SetFormatString(LPCTSTR f)
	{
		m_strFormat = f;
	}
	virtual LPCTSTR GetFormatString() const
	{
		return m_strFormat;
	}

	// Negative input
	virtual void SetNegativeInput(BOOL b = TRUE)
	{
		m_bNegativeInput = b;
	}
	virtual BOOL GetNegativeInput() const
	{
		return m_bNegativeInput;
	}

	// Decimal input
	virtual void SetDecimalInput(BOOL b = TRUE)
	{
		m_bDecimalInput = b;
	}
	virtual BOOL GetDecimalInput() const
	{
		return m_bDecimalInput;
	}

	// Max value
	virtual void SetMaxValue(double d)
	{
		m_dMaxValue = d;
	}
	virtual double GetMaxValue() const
	{
		return m_dMaxValue;
	}

	// Min value
	virtual void SetMinValue(double d)
	{
		m_dMinValue = d;
	}
	virtual double GetMinValue() const
	{
		return m_dMinValue;
	}

protected:
	virtual BOOL IsCharValid(TCHAR& ch);
	virtual BOOL IsNewStringValid(TCHAR& ch);
	virtual BOOL IsPasteCharValid(TCHAR& ch);
	virtual BOOL IsPasteStringValid(TCHAR& ch);

	BOOL IsReadOnly(void) { return !!(GetStyle() & ES_READONLY); }

	virtual BOOL PasteData(const CString& strData);

public:
	virtual void RefreshDisplay();
	virtual void DisplayDefaultValue();

protected:
	CString m_strFormat;		// Format string, the same as printf
	CString m_strPaste;			// The string that gets pasted

	int m_nPastPoint;
	BOOL m_bKeyProcessed;		// Used by OnKeyDown to signal OnChar that a key was handled, don't handle it twice
	BOOL m_bForceProcessing;

	BOOL m_bNegativeInput;		// Can enter negative number
	BOOL m_bDecimalInput;		// Can enter decimal number

	double m_dDefaultValue;		// Value used to replace non legal input
	double m_dMaxValue;			// Max value
	double m_dMinValue;			// Min value

	DECLARE_MESSAGE_MAP();
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg void OnChar(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg LRESULT OnPaste(WPARAM, LPARAM);

public:
#ifdef _DEBUG
	BOOL m_bIsAttached;
	virtual void Dump(CDumpContext& dc) const;
	virtual void AssertValid() const;
#endif

};

#endif // __NUMERICEDIT_H__
