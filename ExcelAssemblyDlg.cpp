
// ExcelAssemblyDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelAssembly.h"
#include "ExcelAssemblyDlg.h"
#include "afxdialogex.h"

#include "CApplication.h"
#include "CRange.h"
#include "CRanges.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
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

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CExcelAssemblyDlg 对话框




CExcelAssemblyDlg::CExcelAssemblyDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcelAssemblyDlg::IDD, pParent)
    , m_sWorksheet(_T(""))
    , m_sReadColumns(_T(""))
    , m_nReadLineFrom(3)
    , m_sColIndex(_T(""))
    , m_sColValue(_T(""))
    , m_sContactColIndex(_T(""))
    , m_sContactColValue(_T(""))
    , m_sContact(_T(""))
    , m_sOutput(_T(""))
    , m_bAppend(FALSE)
    , m_sInput(_T(""))
    , m_bCommonFile(TRUE)
    , m_bContactFile(FALSE)
    , m_bInputSource(TRUE)
    , m_sContactMatchColIndex(_T(""))
    , m_sFileMatchColIndex(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelAssemblyDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Text(pDX, IDC_EDIT_WORKSHEET, m_sWorksheet);
    DDX_Text(pDX, IDC_EDIT_READCOL, m_sReadColumns);
    DDX_Text(pDX, IDC_EDIT_READLINEFROM, m_nReadLineFrom);
    DDX_Text(pDX, IDC_EDIT_COL_INDEX, m_sColIndex);
    DDX_Text(pDX, IDC_EDIT_COL_VALUE, m_sColValue);
    DDX_Text(pDX, IDC_EDIT_CONTACT_COL_INDEX, m_sContactColIndex);
    DDX_Text(pDX, IDC_EDIT_CONTACT_COL_VALUE, m_sContactColValue);
    DDX_Text(pDX, IDC_EDIT_CONTACT, m_sContact);
    DDX_Text(pDX, IDC_EDIT_OUTPUT, m_sOutput);
    DDX_Check(pDX, IDC_CHECK_APPEND, m_bAppend);
    DDX_Text(pDX, IDC_EDIT_INPUT, m_sInput);
    DDX_Check(pDX, IDC_RADIO_COMMONFILE, m_bCommonFile);
    DDX_Check(pDX, IDC_RADIO_CONTACTFILE, m_bContactFile);
    DDX_Check(pDX, IDC_CHECK_INPUTSOURCE, m_bInputSource);
    DDX_Text(pDX, IDC_EDIT_CONTACT_MATCHCOL_INDEX, m_sContactMatchColIndex);
    DDX_Text(pDX, IDC_EDIT_FILE_MATCHCOL_INDEX, m_sFileMatchColIndex);
}

BEGIN_MESSAGE_MAP(CExcelAssemblyDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_COMMAND(IDC_BTN_BROWSE_CONTACT, &CExcelAssemblyDlg::OnBtnBrowseContact)
    ON_COMMAND(IDC_BTN_BROWSE_INPUT, &CExcelAssemblyDlg::OnBtnBrowseInput)
    ON_COMMAND(IDC_BTN_BROWSE_OUTPUT, &CExcelAssemblyDlg::OnBtnBrowseOutput)
    ON_COMMAND(IDC_RADIO_COMMONFILE, &CExcelAssemblyDlg::OnRadioCommonFile)
    ON_COMMAND(IDC_RADIO_CONTACTFILE, &CExcelAssemblyDlg::OnRadioContactFile)
END_MESSAGE_MAP()


// CExcelAssemblyDlg 消息处理程序

BOOL CExcelAssemblyDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
    UpdateData(FALSE);

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CExcelAssemblyDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelAssemblyDlg::OnPaint()
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
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExcelAssemblyDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CExcelAssemblyDlg::OnBtnBrowseContact()
{
    UpdateData(TRUE);
    CFileDialog dlg(TRUE,
        NULL,
        NULL,
        OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
        _T("Excel Files (*.xls;*.xlsx)|*.xls; *.xlsx||"));
    if (dlg.DoModal() == IDOK)
    {
        m_sContact = dlg.GetPathName();
        UpdateData(FALSE);
    }
}

void CExcelAssemblyDlg::OnBtnBrowseInput()
{
    UpdateData(TRUE);
    if (m_bInputSource)
    {
        CFileDialog dlg(TRUE,
            NULL,
            NULL,
            OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
            _T("Excel Files (*.xls;*.xlsx)|*.xls; *.xlsx||"));
        if (dlg.DoModal() == IDOK)
        {
            m_sInput = dlg.GetPathName();
            UpdateData(FALSE);
        }
    }
    else
    {
        TCHAR szPath[MAX_PATH] = {0};  

        ZeroMemory(szPath, sizeof(szPath));   

        BROWSEINFO bi;   
        bi.hwndOwner = m_hWnd;   
        bi.pidlRoot = NULL;   
        bi.pszDisplayName = szPath;   
        bi.lpszTitle = _T("请选择输入源文件所在的文件夹：");   
        bi.ulFlags = 0;   
        bi.lpfn = NULL;   
        bi.lParam = 0;   
        bi.iImage = 0;   
        LPITEMIDLIST lp = SHBrowseForFolder(&bi);   

        if(lp && SHGetPathFromIDList(lp, szPath))   
        {
            m_sInput = szPath;
            UpdateData(FALSE);
        }
    }
}

void CExcelAssemblyDlg::OnBtnBrowseOutput()
{
    UpdateData(TRUE);
    CFileDialog dlg(FALSE,
        _T("xls"),
        NULL,
        OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
        _T("Excel Files (*.xls;*.xlsx)|*.xls; *.xlsx||"));
    if (dlg.DoModal() == IDOK)
    {
        m_sOutput = dlg.GetPathName();
        UpdateData(FALSE);
    }
}

void CExcelAssemblyDlg::OnOK()
{
    if (AfxMessageBox(_T("请最后检查一遍所有设置，点击”确定“开始处理，点击”取消“重新修改设置。"), 
        MB_OKCANCEL|MB_ICONQUESTION) == IDCANCEL)
        return ;

    UpdateData(TRUE);

    if (!CheckInput())
        return ;

    CApplication app;  
    CWorkbooks books;  
    CWorkbook book;  
    CWorksheets sheets;  
    CWorksheet sheet;  
    CRange range;  
    LPDISPATCH lpDisp;      
    COleVariant vResult;  

    CString str;  

    COleVariant  
        covTrue((short)TRUE),  
        covFalse((short)FALSE),  
        covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);  

    //创建Excel 2003服务器(启动Excel)  
    if (!app.CreateDispatch(_T("Excel.Application"),NULL))   
    {   
        AfxMessageBox(_T("Create Excel service failure!"));  
        return;  
    }  

    // 设置为FALSE时，后面的app.Quit();注释要打开  
    // 否则EXCEL.EXE进程会一直存在，并且每操作一次就会多开一个进程  
    app.put_Visible(TRUE);
    books.AttachDispatch(app.get_Workbooks() ,true);  

    //释放对象    
    //range.ReleaseDispatch();  
    //sheet.ReleaseDispatch();  
    //sheets.ReleaseDispatch();  
    //book.ReleaseDispatch();  
    books.ReleaseDispatch();  
    app.ReleaseDispatch();  
}

void CExcelAssemblyDlg::OnRadioCommonFile()
{
    BOOL bChecked = ((CButton*)GetDlgItem(IDC_RADIO_COMMONFILE))->GetCheck();
    GetDlgItem(IDC_EDIT_COL_INDEX)->EnableWindow(bChecked);
    GetDlgItem(IDC_EDIT_COL_VALUE)->EnableWindow(bChecked);
    GetDlgItem(IDC_EDIT_CONTACT_COL_INDEX)->EnableWindow(!bChecked);
    GetDlgItem(IDC_EDIT_CONTACT_COL_VALUE)->EnableWindow(!bChecked);
    GetDlgItem(IDC_EDIT_CONTACT_MATCHCOL_INDEX)->EnableWindow(!bChecked);
    GetDlgItem(IDC_EDIT_FILE_MATCHCOL_INDEX)->EnableWindow(!bChecked);
}

void CExcelAssemblyDlg::OnRadioContactFile()
{
    BOOL bChecked = ((CButton*)GetDlgItem(IDC_RADIO_CONTACTFILE))->GetCheck();
    GetDlgItem(IDC_EDIT_COL_INDEX)->EnableWindow(!bChecked);
    GetDlgItem(IDC_EDIT_COL_VALUE)->EnableWindow(!bChecked);
    GetDlgItem(IDC_EDIT_CONTACT_COL_INDEX)->EnableWindow(bChecked);
    GetDlgItem(IDC_EDIT_CONTACT_COL_VALUE)->EnableWindow(bChecked);
    GetDlgItem(IDC_EDIT_CONTACT_MATCHCOL_INDEX)->EnableWindow(bChecked);
    GetDlgItem(IDC_EDIT_FILE_MATCHCOL_INDEX)->EnableWindow(bChecked);
}

BOOL CExcelAssemblyDlg::CheckInput()
{
    if(!::PathFileExists(m_sInput))
    {
        AfxMessageBox(_T("输入源文件或文件夹不存在，请检查后重新输入！"), MB_OK|MB_ICONSTOP);
        GetDlgItem(IDC_EDIT_INPUT)->SetFocus();
        return FALSE;
    }

    if (m_bInputSource && PathIsDirectory(m_sInput))
    {
        AfxMessageBox(_T("输入源类型选择了“文件”，输入源路径却指向了一个文件夹，请检查后重新输入！"), MB_OK|MB_ICONSTOP);
        return FALSE;
    }

    if (!m_bInputSource && !PathIsDirectory(m_sInput))
    {
        AfxMessageBox(_T("输入源类型选择了“文件夹”，输入源路径却指向了一个文件，请检查后重新输入！"), MB_OK|MB_ICONSTOP);
        return FALSE;
    }

    if (m_sWorksheet.IsEmpty())
    {
        AfxMessageBox(_T("请输入要读取的工作表的序号或名称。"), MB_OK|MB_ICONSTOP);
        GetDlgItem(IDC_EDIT_WORKSHEET)->SetFocus();
        return FALSE;
    }

    if (m_sReadColumns.IsEmpty())
    {
        AfxMessageBox(_T("请输入要汇总的列号。"));
        GetDlgItem(IDC_EDIT_READCOL)->SetFocus();
        return FALSE;
    }

    CString readCol = m_sReadColumns.MakeUpper();
    for(int i = 0; i < readCol.GetLength(); i++)
    {
        if (readCol.GetAt(i) != ' ' && (readCol.GetAt(i) < 'A' || readCol.GetAt(i) > 'Z'))
        {
            AfxMessageBox(_T("汇总的列号只能输入字母，以空格分隔。"), MB_OK|MB_ICONSTOP);
            GetDlgItem(IDC_EDIT_READCOL)->SetFocus();
            return FALSE;
        }
    }

    if (m_nReadLineFrom > 10)
    {
        if (AfxMessageBox(_T("确定要从这么大的行数开始读取吗？"), MB_YESNO|MB_ICONQUESTION) == IDNO)
        {
            GetDlgItem(IDC_EDIT_READLINEFROM)->SetFocus();
            return FALSE;
        }
    }

    if (m_bCommonFile)
    {
        if (m_sColIndex.IsEmpty())
        {
            AfxMessageBox(_T("请输入汇总条件中本文件的列号。"), MB_OK|MB_ICONSTOP);
            GetDlgItem(IDC_EDIT_COL_INDEX)->SetFocus();
            return FALSE;
        }

        CString col = m_sColIndex.MakeUpper();
        for(int i = 0; i < col.GetLength(); i++)
        {
            if (col.GetAt(i) < 'A' || col.GetAt(i) > 'Z')
            {
                AfxMessageBox(_T("汇总条件中本文件的列号只能输入字母。"), MB_OK|MB_ICONSTOP);
                GetDlgItem(IDC_EDIT_COL_INDEX)->SetFocus();
                return FALSE;
            }
        }

        if (m_sColValue.IsEmpty())
        {
            AfxMessageBox(_T("汇总条件中本文件的列匹配的值没有填写！"), MB_OK|MB_ICONSTOP);
            GetDlgItem(IDC_EDIT_COL_VALUE)->SetFocus();
            return FALSE;
        }
    }

    if (m_bContactFile)
    {
        if (m_sContactColIndex.IsEmpty() || m_sContactMatchColIndex.IsEmpty() || m_sFileMatchColIndex.IsEmpty())
        {
            AfxMessageBox(_T("请输入汇总条件中花名册和源文件的列号。"), MB_OK|MB_ICONSTOP);
            GetDlgItem(IDC_EDIT_CONTACT_COL_INDEX)->SetFocus();
            return FALSE;
        }

        CString col = m_sContactColIndex.MakeUpper();
        for(int i = 0; i < col.GetLength(); i++)
        { 
            if (col.GetAt(i) < 'A' || col.GetAt(i) > 'Z')
            {
                AfxMessageBox(_T("汇总条件中花名册的列号只能输入字母。"), MB_OK|MB_ICONSTOP);
                GetDlgItem(IDC_EDIT_CONTACT_COL_INDEX)->SetFocus();
                return FALSE;
            }
        }

        col = m_sContactMatchColIndex.MakeUpper();
        for(int i = 0; i < col.GetLength(); i++)
        { 
            if (col.GetAt(i) < 'A' || col.GetAt(i) > 'Z')
            {
                AfxMessageBox(_T("汇总条件中花名册待匹配的列号只能输入字母。"), MB_OK|MB_ICONSTOP);
                GetDlgItem(IDC_EDIT_CONTACT_COL_INDEX)->SetFocus();
                return FALSE;
            }
        }

        col = m_sFileMatchColIndex.MakeUpper();
        for(int i = 0; i < col.GetLength(); i++)
        { 
            if (col.GetAt(i) < 'A' || col.GetAt(i) > 'Z')
            {
                AfxMessageBox(_T("汇总条件中源文件待匹配的列号只能输入字母。"), MB_OK|MB_ICONSTOP);
                GetDlgItem(IDC_EDIT_CONTACT_COL_INDEX)->SetFocus();
                return FALSE;
            }
        }

        if (m_sContactColValue.IsEmpty())
        {
            AfxMessageBox(_T("汇总条件中花名册的列匹配的值没有填写！"), MB_OK|MB_ICONSTOP);
            GetDlgItem(IDC_EDIT_CONTACT_COL_VALUE)->SetFocus();
            return FALSE;
        }

        if (m_sContact.IsEmpty())
        {
            AfxMessageBox(_T("你在汇总条件中选择了使用花名册，却没有指定花名册的路径。"), MB_OK|MB_ICONSTOP);
            GetDlgItem(IDC_EDIT_CONTACT)->SetFocus();
            return FALSE;
        }
    }

    if (m_sOutput.IsEmpty())
    {
        AfxMessageBox(_T("请填写输出文件路径。"), MB_OK|MB_ICONSTOP);
        GetDlgItem(IDC_EDIT_OUTPUT)->SetFocus();
        return FALSE;
    }

    if (!m_bAppend && PathFileExists(m_sOutput))
    {
        if (AfxMessageBox(_T("输出文件已存在，确定要覆盖吗？"), MB_YESNO | MB_ICONQUESTION) == IDNO)
        {
            GetDlgItem(IDC_EDIT_OUTPUT)->SetFocus();
            return FALSE;
        }
    }

    return TRUE;
}
