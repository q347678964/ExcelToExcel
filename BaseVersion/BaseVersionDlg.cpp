
// BaseVersionDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "BaseVersion.h"
#include "BaseVersionDlg.h"
#include "afxdialogex.h"
#include "Config.h"
#include "FormatChange.h"
#include "Example.h"
#include "ExcelHandler.h"
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


// CBaseVersionDlg �Ի���




CBaseVersionDlg::CBaseVersionDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CBaseVersionDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);
}

void CBaseVersionDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CBaseVersionDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BUTTON_Start, &CBaseVersionDlg::OnBnClickedButtonStart)
END_MESSAGE_MAP()


// CBaseVersionDlg ��Ϣ��������

BOOL CBaseVersionDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	//m_hIcon = AfxGetApp()->LoadIcon(IDI_ICON1);
	//SetIcon(m_hIcon, TRUE); // Set big icon  
	//SetIcon(m_hIcon, FALSE); // Set small icon; 

	// ��������...���˵������ӵ�ϵͳ�˵��С�

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
	// TODO: �ڴ����Ӷ���ĳ�ʼ������
	//this->Printf((CString)("[Dialog]��ʼ���������!\r\n"));
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CBaseVersionDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void CBaseVersionDlg::MoveTo(CWnd **SrcWnd,CWnd **DestWnd,int DestID,DEST_TYPE DestType,MOVE_DIR MoveMethod){
	CRect RectTemp;

	switch(MoveMethod){
		case MOVE_RIGHT:
			switch(DestType){
				case DEST_EDIT:
					(*SrcWnd)->GetWindowRect(RectTemp);//��ȡĿ������Ļ�ϵ����꣬��Ҫת������������
					ScreenToClient(RectTemp);
					(*DestWnd) = GetDlgItem(DestID);
					(*DestWnd)->SetWindowPos( NULL,RectTemp.right+10,RectTemp.top,0,0,SWP_NOZORDER|SWP_NOSIZE);	//·�����ڣ�ֻ�ı����꣬���ı��С
					break;
				case DEST_BUTTON:
					(*SrcWnd)->GetWindowRect(RectTemp);//��ȡĿ������Ļ�ϵ����꣬��Ҫת������������
					ScreenToClient(RectTemp);
					(*DestWnd) = GetDlgItem(DestID);
					(*DestWnd)->SetWindowPos( NULL,RectTemp.right+10,RectTemp.top,0,0,SWP_NOZORDER|SWP_NOSIZE);	//·���򿪰�ť��ֻ�ı����꣬���ı��С
					((CButton *)GetDlgItem(DestID))->SetIcon(AfxGetApp()->LoadIcon(IDI_ICON1));			//��ťͼƬ
					//��Ч��������AfxGetApp()->LoadIcon(IDI_ICON1) ����icon
					//((CMFCButton *)GetDlgItem(IDC_BUTTON_OpenFile)) ��ȡ��ť���
					//SetIcon ��ť��������
					break;
			}
			break;
		case MOVE_BOTTOM:
			switch(DestType){
				case DEST_EDIT:
					(*SrcWnd)->GetWindowRect(RectTemp);//��ȡĿ������Ļ�ϵ����꣬��Ҫת������������
					ScreenToClient(RectTemp);
					(*DestWnd) = GetDlgItem(DestID);
					(*DestWnd)->SetWindowPos( NULL,RectTemp.left,RectTemp.bottom+10,0,0,SWP_NOZORDER|SWP_NOSIZE);	//·�����ڣ�ֻ�ı����꣬���ı��С
					break;
				case DEST_BUTTON:
					(*SrcWnd)->GetWindowRect(RectTemp);//��ȡĿ������Ļ�ϵ����꣬��Ҫת������������
					ScreenToClient(RectTemp);
					(*DestWnd) = GetDlgItem(DestID);
					(*DestWnd)->SetWindowPos( NULL,RectTemp.left,RectTemp.bottom+10,0,0,SWP_NOZORDER|SWP_NOSIZE);	//·�����ڣ�ֻ�ı����꣬���ı��С
					((CButton *)GetDlgItem(DestID))->SetIcon(AfxGetApp()->LoadIcon(IDI_ICON1));			//��ťͼƬ
					break;
			}
			break;
		default:
			break;
	}
}
/*
[0] = ���Դ���EDIT
[1] = ·��EDIT
[2] = ���ļ�BUTTON
[3] = ������PROCESS
[4] = ����BUTTON
[5] = Picture Control
*/
void CBaseVersionDlg::DlgPaintInit(void)
{
	CImage mImage;  
    if(mImage.Load(_T(CFG_CSTRING_BGP)) == S_OK)  {
		CWnd *pWnd[20];  
		CRect RectTemp;
        //�����ô��ڱ��ֺͱ���ͼһ�� 
		int WinDlgWidth = mImage.GetWidth();
		int WinDlgHeight = mImage.GetHeight();
		SetWindowPos(NULL,0,0,WinDlgWidth,WinDlgHeight,SWP_NOMOVE);
		
		pWnd[0] = GetDlgItem(IDC_EDIT_Debug);
		pWnd[0]->SetWindowPos( NULL,10,10,WinDlgWidth - 10,WinDlgHeight - 200,SWP_NOZORDER);//���Դ���,���ݴ����С���任
		
		this->MoveTo(&pWnd[0],&pWnd[1],IDC_BUTTON_Start,DEST_BUTTON,MOVE_BOTTOM);
		//mImage.Draw(GetDC()->GetSafeHdc(),CRect(0,0,WinDlgWidth,WinDlgHeight));//���������������ַ������ƣ�����˸
		
		{	//��������
			CBitmap	bmpBackground;		//����һ��λͼ���
			FormatChange FC;
			FC.CImage2CBitmap(mImage,bmpBackground);
			//bmpBackground.LoadBitmap(IDB_BITMAP1);   //����ͼƬ��λͼ���	

			CRect   WinDlg;   
			GetClientRect(&WinDlg);			//��ȡ����Ĵ�С

			CDC *BGPDCMem = new CDC;;		//�����ڴ�ͼƬCDC
			CPaintDC WinDlgDc(this);					//��δ����������ñ���ͼ����ʼ��DC���ƶ���Ϊ���屾��
			BGPDCMem->CreateCompatibleDC(&WinDlgDc);   //����һ������ʾ���豸���ݼ��ݵ��ڴ��豸����
			BGPDCMem->SetBkMode(TRANSPARENT);
			BGPDCMem->SelectObject(&bmpBackground); //��SelectObject��λͼѡ���ڴ��豸����  
#if 0	//���췽ʽ����
			BITMAP   bitmap;	//BITMAP�ṹ���ڴ��λͼ��Ϣ
			bmpBackground.GetBitmap(&bitmap);	//��ͼƬ�л�ȡͼƬ�Ŀ��ߵ�bitmap
			WinDlgDc.StretchBlt(0,0,WinDlg.Width(),WinDlg.Height(),&g_BGPDCMem,0,0,bitmap.bmWidth,bitmap.bmHeight,SRCCOPY); //��DC�ڵ�ͼƬ��������֮��PO������,Stretch:����
#else
			WinDlgDc.BitBlt(0,0,WinDlg.Width(),WinDlg.Height(),BGPDCMem,0,0,SRCCOPY);	//���ڴ��ͼƬpo��������
#endif
			delete BGPDCMem;
		}
	}



}

void CBaseVersionDlg::OnPaint()
{
	static int FirstIn = 1;
	if(FirstIn){
		FirstIn = 0;
		this->DlgPaintInit();
	}
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
		CDialog::OnPaint();
	}

	//this->Printf((CString)("[Dialog]���»������!\r\n"));
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CBaseVersionDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//�������� 
BOOL CBaseVersionDlg::PreTranslateMessage(MSG* pMsg)  
{  
    // TODO:  �ڴ�����ר�ô����/����û���  
    if (pMsg->message == WM_KEYDOWN)  
    {  
        switch (pMsg->wParam)  
        {  
        case'I':  
            //if (::GetKeyState(VK_CONTROL) < 0)//�����Shift+X�����  
                //�ĳ�VK_SHIFT  
                MessageBox(_T("Hello"));  
            return TRUE;  
        }  

		Example Ex;
		Ex.DlgMsgListen(pMsg->wParam);
    }  
    return CDialog::PreTranslateMessage(pMsg);  
}  

HBRUSH CBaseVersionDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{ 
	HBRUSH hbr = CDialog::OnCtlColor(pDC, pWnd, nCtlColor); 
	 switch (pWnd->GetDlgCtrlID()){
		case IDC_EDIT_Debug:
			pDC->SetTextColor(EDIT_PRINT_TEXT_RGB); //����������ɫ
			pDC->SetBkMode(TRANSPARENT);	//�������屳��Ϊ͸������������ֱ�Ӵ���Edit����ɫ
			return (HBRUSH)CreateSolidBrush(EDIT_PRINT_BG_RGB);// ���ñ���ɫ��ˢ
			break;
	 }

	// TODO: Change any attributes of the DC here
	if (nCtlColor==CTLCOLOR_EDIT)//�����ǰ�ؼ������ı�
	{ 

	}else if(nCtlColor==CTLCOLOR_STATIC){	//STATIC PICTURE

	}
	else if (nCtlColor==CTLCOLOR_BTN) //�����ǰ�ؼ����ڰ�ť
	{ 

	} 
	// TODO: Return a different brush if the default is not desired
	return hbr; 
}


BOOL CBaseVersionDlg::Printf(CString string){
	static CString DebugCStringAll;
	DebugCStringAll += string;
	SetDlgItemText(IDC_EDIT_Debug,DebugCStringAll);
	CEdit *DebugEdit = (CEdit*)GetDlgItem(IDC_EDIT_Debug);
	DebugEdit->LineScroll(DebugEdit->GetLineCount()-1,0); 
	return 0;
}


void CBaseVersionDlg::OnBnClickedButtonStart()
{
	CExcelHandler ExcelHdlr;

	ExcelHdlr.Excel_AllHandler();
}