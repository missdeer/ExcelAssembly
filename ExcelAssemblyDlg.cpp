
// ExcelAssemblyDlg.cpp : ʵ���ļ�
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


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
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


// CExcelAssemblyDlg �Ի���




CExcelAssemblyDlg::CExcelAssemblyDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CExcelAssemblyDlg::IDD, pParent)
    , m_bSingleFileCheck(TRUE)
    , m_bFolderCheck(FALSE)
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
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelAssemblyDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_Check(pDX, IDC_RADIO_FILE, m_bSingleFileCheck);
    DDX_Check(pDX, IDC_RADIO_FOLDER, m_bFolderCheck);
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
}

BEGIN_MESSAGE_MAP(CExcelAssemblyDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
    ON_WM_QUERYDRAGICON()
    ON_COMMAND(IDC_BTN_BROWSE_CONTACT, &CExcelAssemblyDlg::OnBtnBrowseContact)
    ON_COMMAND(IDC_BTN_BROWSE_INPUT, &CExcelAssemblyDlg::OnBtnBrowseInput)
    ON_COMMAND(IDC_BTN_BROWSE_OUTPUT, &CExcelAssemblyDlg::OnBtnBrowseOutput)
END_MESSAGE_MAP()


// CExcelAssemblyDlg ��Ϣ�������

BOOL CExcelAssemblyDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
    UpdateData(FALSE);

	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
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

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CExcelAssemblyDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
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
    if (m_bSingleFileCheck)
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
        bi.lpszTitle = _T("��ѡ������Դ�ļ����ڵ��ļ��У�");   
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

    //����Excel 2003������(����Excel)  
    if (!app.CreateDispatch(_T("Excel.Application"),NULL))   
    {   
        AfxMessageBox(_T("Create Excel service failure!"));  
        return;  
    }  

    // ����ΪFALSEʱ�������app.Quit();ע��Ҫ��  
    // ����EXCEL.EXE���̻�һֱ���ڣ�����ÿ����һ�ξͻ�࿪һ������  
    app.put_Visible(TRUE);
    books.AttachDispatch(app.get_Workbooks() ,true);  

    //�ͷŶ���    
    //range.ReleaseDispatch();  
    //sheet.ReleaseDispatch();  
    //sheets.ReleaseDispatch();  
    //book.ReleaseDispatch();  
    books.ReleaseDispatch();  
    app.ReleaseDispatch();  
}

