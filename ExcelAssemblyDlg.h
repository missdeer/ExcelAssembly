
// ExcelAssemblyDlg.h : ͷ�ļ�
//

#pragma once


// CExcelAssemblyDlg �Ի���
class CExcelAssemblyDlg : public CDialogEx
{
// ����
public:
	CExcelAssemblyDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EXCELASSEMBLY_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
