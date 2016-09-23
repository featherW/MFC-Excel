// ExcelClassDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelClass.h"
#include "ExcelClassDlg.h"
#include "CExcelOperation.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CExcelClassDlg 对话框




CExcelClassDlg::CExcelClassDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CExcelClassDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelClassDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CExcelClassDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, &CExcelClassDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CExcelClassDlg 消息处理程序

BOOL CExcelClassDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CExcelClassDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelClassDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExcelClassDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void MyCallBack(void *hFun)
{
	CExcelOperation *pExcelOper= (CExcelOperation*)hFun;
	BYTE i = 0;
	BYTE j = 0;
	CString csPortTmp;
	long lMaxRowNum = pExcelOper->getMaxRowNum();
	long lMaxColNum = pExcelOper->getMaxColNum();
	CRange rangeBasicCells = pExcelOper->getCell();
	CRange rangeTempCells;
	COleVariant covResult;

	for (i = 1; i != lMaxRowNum+1; i++) 
	{
		for (j = 1; j != lMaxColNum+1; j++)
		{
			rangeTempCells.AttachDispatch(rangeBasicCells.get_Item(COleVariant(i), COleVariant(j)).pdispVal);
			covResult = rangeTempCells.get_Value2();
			if (covResult.vt != VT_BSTR)
			{
				continue;
			}
			csPortTmp = covResult.bstrVal;
			MessageBox(NULL,csPortTmp.GetBuffer(0),NULL,NULL);
		}
	}
}

void CExcelClassDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	CString csXlsPath;
	CFileDialog cfd(true,
		_T(".xls"),
		NULL,
		OFN_FILEMUSTEXIST|OFN_HIDEREADONLY,
		_T("Excel file   (*.xls; *.xlsx)|*.xls; *.xlsx|All   Files   (*.*)|*.*||"),
		this); 
	if(cfd.DoModal()!=IDOK)
	{ 
		return;
	}
	csXlsPath = cfd.GetPathName();
	if(csXlsPath == "")
	{
		return;
	}
	CExcelOperation excelApp;
	string strErrMsg;
	excelApp.openExcelFile(csXlsPath);
	excelApp.setDealCB(MyCallBack);
	excelApp.startDealExcel();
	strErrMsg = excelApp.getErrMsg();
	MessageBox(strErrMsg.c_str());
	excelApp.closeExcelFile();
}

