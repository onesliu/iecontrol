
// testieDlg.h : header file
//

#pragma once
#include "afxwin.h"


// CtestieDlg dialog
class CtestieDlg : public CDialogEx
{
// Construction
public:
	CtestieDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_TESTIE_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;
	CComQIPtr<IWebBrowser2> spBrowser;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
private:
	int findFromShell(void);
public:
	afx_msg void OnBnClickedOk();
public:
	CStatic m_info;
};
