
// controlcliDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ControlCli.h"
#include "ControlCliDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CCtlPanelDlg dialog




CCtlPanelDlg::CCtlPanelDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CCtlPanelDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CCtlPanelDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_INFO, m_info);
}

BEGIN_MESSAGE_MAP(CCtlPanelDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDOK, &CCtlPanelDlg::OnBnClickedOk)
	ON_WM_TIMER()
	ON_NOTIFY(NM_CLICK, IDC_SYSLINK1, &CCtlPanelDlg::OnNMClickSyslink1)
	ON_NOTIFY(NM_CLICK, IDC_SYSLINK2, &CCtlPanelDlg::OnNMClickSyslink1)
END_MESSAGE_MAP()


// CCtlPanelDlg message handlers

BOOL CCtlPanelDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here
	CoInitialize(NULL);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CCtlPanelDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CCtlPanelDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CCtlPanelDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


int CCtlPanelDlg::findFromShell(void)
{
	CComPtr<IShellWindows> spShellWin;
	HRESULT hr = spShellWin.CoCreateInstance(CLSID_ShellWindows);
	if (FAILED(hr)) return -1;

	long count;
	CComPtr<IDispatch> spDisp;

	spShellWin->get_Count(&count);
	for (int i = 0; i < count; i++) {
		hr = spShellWin->Item(CComVariant(i), &spDisp);
		if (FAILED(hr)) continue;

		spBrowser = spDisp;
		if (!spBrowser) continue;
		spDisp.Release();

		CComBSTR url;
		spBrowser->get_LocationURL(&url);

		CUrl psUrl;
		psUrl.CrackUrl(url);

		if (CString(psUrl.GetUrlPath()) != L"/") continue;

		CString csHost(psUrl.GetHostName());
		if (csHost.Find(L"www.google") == -1)
			continue;

		hr = spBrowser->get_Document(&spDisp);
		if (FAILED(hr)) continue;

		CComQIPtr<IHTMLDocument2> spDoc = spDisp;
		if (!spDoc) continue;
		spDisp.Release();
		
		CComQIPtr<IHTMLElementCollection> spColls;
		spDoc->get_forms(&spColls);
		
		spColls->item(CComVariant(), CComVariant(0), &spDisp);
		CComQIPtr<IHTMLFormElement> spForm(spDisp);
		if (!spForm) continue;
		spDisp.Release();
		
		long fcount;
		spForm->get_length(&fcount);
		CComQIPtr<IHTMLInputTextElement> spText;
		for(int j = 0; j < fcount; j++) {
			
			CComDispatchDriver spInputElement;
			hr = spForm->item(CComVariant(j), CComVariant(), &spInputElement);
			if (FAILED(hr)) continue;

			CComVariant vName,vVal,vType;
			spInputElement.GetPropertyByName(L"name", &vName);
			spInputElement.GetPropertyByName(L"value", &vVal);
			spInputElement.GetPropertyByName(L"type", &vType);
			
			if (CString(vType) == L"text") {
				spText = spInputElement;
			}
		}

		CComVariant vStr(L"Merry Christmas");
		hr = spText->put_value(vStr.bstrVal);
		spForm->submit();
		return 0;
	}
	
	return -1;
}

void CCtlPanelDlg::OnBnClickedOk()
{
	// TODO: 在此添加控件通知处理程序代码
	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	ZeroMemory( &si, sizeof(si) );
	ZeroMemory( &pi, sizeof(pi) );
	si.cb = sizeof(si);
	TCHAR  pszCmdLine[] = L"c:\\program files\\internet explorer\\iexplore.exe www.google.com";
	if( CreateProcess( NULL, pszCmdLine, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi) == 0 ) {
			m_info.SetWindowTextW(L"Open a browser error!");
			UpdateData();
			return;
	}

	if ( WaitForInputIdle(pi.hProcess, 10000) != 0 ) {
			m_info.SetWindowTextW(L"Open a browser error!");
			UpdateData();
			return;
	}

	CloseHandle(pi.hThread);
	CloseHandle(pi.hProcess);

	SetTimer(TIMER1, 200, 0);
	//CDialogEx::OnOK();
}


void CCtlPanelDlg::OnTimer(UINT_PTR nIDEvent)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	if ( findFromShell() == 0 ) {
		m_info.SetWindowTextW(L"All done. Merry Christmas!");
		KillTimer(TIMER1);
	}
	else {
		m_info.SetWindowTextW(L"Browser is opened, getting google page...");
	}

	UpdateData();

	CDialogEx::OnTimer(nIDEvent);
}


void CCtlPanelDlg::OnNMClickSyslink1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// TODO: 在此添加控件通知处理程序代码
	PNMLINK pNMLink = (PNMLINK) pNMHDR;
	ShellExecuteW(NULL, L"open", pNMLink->item.szUrl, NULL, NULL, SW_SHOWNORMAL);
	*pResult = 0;
}
