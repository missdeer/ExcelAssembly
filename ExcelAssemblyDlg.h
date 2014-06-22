
// ExcelAssemblyDlg.h : 头文件
//

#pragma once


// CExcelAssemblyDlg 对话框
class CExcelAssemblyDlg : public CDialogEx
{
// 构造
public:
	CExcelAssemblyDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EXCELASSEMBLY_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
    afx_msg void OnBtnBrowseContact();
    afx_msg void OnBtnBrowseInput();
    afx_msg void OnBtnBrowseOutput();
    afx_msg void OnOK();
    afx_msg void OnRadioCommonFile();
    afx_msg void OnRadioContactFile();
	DECLARE_MESSAGE_MAP()
public:
    CString m_sWorksheet;
    CString m_sReadColumns;
    int m_nReadLineFrom;
    CString m_sColIndex;
    CString m_sColValue;
    CString m_sContactColIndex;
    CString m_sContactColValue;
    CString m_sContact;
    CString m_sOutput;
    BOOL m_bAppend;
    CString m_sInput;
    BOOL m_bCommonFile;
    BOOL m_bContactFile;
    BOOL m_bInputSource;
};
