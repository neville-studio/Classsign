// Classsign.cpp : ����Ӧ�ó������ڵ㡣
/*
	Copyright by Xupeng Studio,all rights reserved.
	This code start on Aug.1,2021.


*/

#include "stdafx.h"
#include "Classsign.h"
#include "windowsx.h"
#include "commdlg.h"
#include "commctrl.h"
#include "io.h"
#include "vector"
#include <iostream>
#include "AES.h"
#import <msxml6.dll> raw_interface_only
#include "msxml.h"
#include "Shlobj.h"
#include <fstream>
#include "time.h"

#pragma comment(linker,"\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#define MAX_LOADSTRING 100
#define base64 L"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz+="
using namespace MSXML2;
using namespace std;
HWND hWndMain;
HWND AdminTab[4];
INT_PTR CALLBACK Admin(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);


// ȫ�ֱ���:
HINSTANCE hInst;								// ��ǰʵ��
TCHAR szTitle[MAX_LOADSTRING];					// �������ı�
TCHAR szWindowClass[MAX_LOADSTRING];			// ����������
bool start=true;
OPENFILENAME classignmainfilename;
HANDLE hf;              // file handle
int option;
/*
	option�����ʽ˵����option = 1ʱ����ʾ��������ǩ��
	option = 2ʱ����ʾ��������Ա��������
*/
INT_PTR CALLBACK Userlog(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
// �˴���ģ���а����ĺ�����ǰ������:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    Start(HWND hDlg, UINT message,WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK	Password(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK	UserProcess(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK	SetPassword(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK	SetNames(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);

INT_PTR CALLBACK	Booking(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK	Advance(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK	SignFile(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK	Student(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam);
CHAR* pluszero(int num);
//vector <vector <CHAR>> UserInformation;
//vector <vector <CHAR>> UserPassword;
//vector <vector <CHAR>> UserID;
int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{

	WIN32_FIND_DATA temp;
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

 	// TODO: �ڴ˷��ô��롣
	MSG msg;
	HACCEL hAccelTable;

	// ��ʼ��ȫ���ַ���
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_CLASSSIGN, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);
	hInst = hInstance;
	temp.cFileName[0] = '\0';
	// ִ��Ӧ�ó����ʼ��:
	TCHAR szBuffer[4096];
	GetEnvironmentVariable(_T("APPDATA"), szBuffer, 2048);
	TCHAR szBuffer2[2048] = L"\\Xupeng Studio\\Classsign\\Userprof.xupestd";
	wcscat_s(szBuffer, szBuffer2);
	FindFirstFile(szBuffer, &temp);
	if (temp.cFileName[0] == '\0') {
		DialogBox(hInstance, MAKEINTRESOURCE(FIRSTRUN_STEP1), NULL, SetPassword);
		start = false;
	}
	else 
	{
		DialogBox(hInstance, MAKEINTRESOURCE(IDD_STARTDIALOG), NULL, Start);
	}
	

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_CLASSSIGN));

	// ����Ϣѭ��:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return (int) msg.wParam;
}



//
//  ����: MyRegisterClass()
//
//  Ŀ��: ע�ᴰ���ࡣ
//
//  ע��:
//
//    ����ϣ��
//    �˴�������ӵ� Windows 95 �еġ�RegisterClassEx��
//    ����֮ǰ�� Win32 ϵͳ����ʱ������Ҫ�˺��������÷������ô˺���ʮ����Ҫ��
//    ����Ӧ�ó���Ϳ��Ի�ù�����
//    ����ʽ��ȷ�ġ�Сͼ�ꡣ
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;


	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_CLASSSIGN));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_CLASSSIGN);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

	return RegisterClassEx(&wcex);
}

//
//   ����: InitInstance(HINSTANCE, int)
//
//   Ŀ��: ����ʵ�����������������
//
//   ע��:
//
//        �ڴ˺����У�������ȫ�ֱ����б���ʵ�������
//        ��������ʾ�����򴰿ڡ�
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   HWND hWnd;
   hInst = hInstance; // ��ʵ������洢��ȫ�ֱ�����
   
	RECT rect;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &rect, 0);
	int cx = rect.right - rect.left;
	int cy = rect.bottom - rect.top;
	hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
		cx / 3, 0, cx / 3, cy, NULL, NULL, hInstance, NULL);
	hWndMain = hWnd;
	
	
   if (!hWnd)
   {
      return FALSE;
   }
   ShowWindow(hWndMain, SW_SHOW);
	UpdateWindow(hWndMain);
   return TRUE;
}

//
//  ����: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  Ŀ��: ���������ڵ���Ϣ��
//
//  WM_COMMAND	- ����Ӧ�ó���˵�
//  WM_PAINT	- ����������
//  WM_DESTROY	- �����˳���Ϣ������
//
//ǩ������Ĵ�������
HWND MAIN_LIST1,MAIN_RECORDSLIST,MAIN_STATIC1,MAIN_NAMECOMBO,MAIN_SIGNOPTION,MAIN_PASSWORDEDIT,MAIN_BUTTON1,MAIN_BUTTON2,MAIN_STATIC2,MAIN_STATIC3;
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	int formheight,formwidth;
	int xoff,yoff,itemCount;
	PAINTSTRUCT ps;
	TCHAR szFile[1024];
	HDC hdc;
	HFONT MainFont,List1Font;
	int FormWidth,ClientWidth,ClientHeight,FormHeight,cx,cy ;
	RECT rect;
	MSXML2::IXMLDOMDocumentPtr XMLDOC;
	MSXML2::IXMLDOMElementPtr XMLROOT;
	MSXML2::IXMLDOMElementPtr XMLELEMENT;
	MSXML2::IXMLDOMNodeListPtr XMLNODES; //ĳ���ڵ�������ֽڵ�
	MSXML2::IXMLDOMNamedNodeMapPtr XMLNODEATTS;//ĳ���ڵ����������;
	MSXML2::IXMLDOMNodePtr XMLNODE;
	LV_ITEM lvi;
	LV_COLUMN COl;
	HRESULT XmlFile;
	switch (message)
	{
	case WM_CREATE:
		//���ڻ���
		MainFont = CreateFont(13,0,0,0,0,FALSE,FALSE,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,DEFAULT_PITCH|FF_SWISS,_T("����"));
		List1Font = CreateFont(20,0,0,0,0,FALSE,FALSE,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,DEFAULT_PITCH|FF_SWISS,_T("����"));
		RECT MainFormRect;
		RECT MainFormClientRect;
		GetClientRect(hWnd,&MainFormClientRect);
		GetWindowRect(hWnd,&MainFormRect);
		FormWidth = MainFormRect.right - MainFormRect.left;
		FormHeight = MainFormRect.bottom - MainFormRect.top;
		ClientWidth = MainFormClientRect.right - MainFormClientRect.left;
		ClientHeight = MainFormClientRect.bottom - MainFormClientRect.top;
		MAIN_LIST1=CreateWindow(_T("LISTBOX"),_T("û��ǩ����ͬ־"),WS_CHILD|WS_VISIBLE|LBS_SORT|LBS_STANDARD|WS_HSCROLL|LBS_DISABLENOSCROLL,0,25,(ClientWidth - 10)/3,ClientHeight - 25,hWnd,(HMENU)1,hInst,NULL);
		MAIN_RECORDSLIST=CreateWindowEx(WS_EX_CLIENTEDGE,WC_LISTVIEW, _T("ǩ����¼"),WS_CHILD|WS_VISIBLE |LVS_REPORT | WS_TABSTOP,(ClientWidth - 10)/3*2+10,25,(ClientWidth-10)/3,ClientHeight - 25,hWnd,(HMENU)2,hInst,NULL);
		MAIN_STATIC1 = CreateWindow(_T("STATIC"),_T("������ûǩ����������ѡ�񲢽���ǩ����֤��"),WS_CHILD | WS_VISIBLE,0,0,(ClientWidth - 10)/3,25,hWnd,(HMENU)3,hInst,NULL);
		MAIN_STATIC2 = CreateWindow(_T("STATIC"),_T("ѡ����֤��ʽ��"),WS_CHILD | WS_VISIBLE,(ClientWidth - 10)/3+5,35,(ClientWidth - 10)/3,15,hWnd,(HMENU)3,hInst,NULL);
		MAIN_BUTTON1 = CreateWindow(_T("BUTTON"),_T("����ǩ��"),WS_CHILD | WS_VISIBLE| BS_AUTORADIOBUTTON ,(ClientWidth - 10)/3+5,50,(ClientWidth - 10)/3,15,hWnd,(HMENU)4,hInst,NULL);
		MAIN_BUTTON2 = CreateWindow(_T("BUTTON"),_T("�����û�����ǩ����"),WS_CHILD | WS_VISIBLE|BS_AUTORADIOBUTTON ,(ClientWidth - 10)/3+5,65,(ClientWidth - 10)/3,15,hWnd,(HMENU)5,hInst,NULL);
		MAIN_SIGNOPTION = CreateWindow(_T("BUTTON"),_T("ǩ��"),WS_CHILD | WS_VISIBLE,(ClientWidth - 10)/3+5,100,(ClientWidth - 10)/3,25,hWnd,(HMENU)7,hInst,NULL);
		MAIN_NAMECOMBO = CreateWindow(_T("COMBOBOX"),_T("��ѡ��"),WS_CHILD | WS_VISIBLE|CBS_DROPDOWN|CBS_AUTOHSCROLL|CBS_SORT,(ClientWidth - 10)/3+5,0,(ClientWidth - 10)/3,25,hWnd,(HMENU)6,hInst,NULL);

		MAIN_PASSWORDEDIT = CreateWindow(_T("EDIT"),_T(""),WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER,(ClientWidth - 10)/3+5,80,(ClientWidth - 10)/3,20,hWnd,(HMENU)8,hInst,NULL);
		MAIN_STATIC3 = CreateWindow(_T("STATIC"),_T("ǩ����¼"),WS_CHILD | WS_VISIBLE,(ClientWidth - 10)/3*2+10,0,(ClientWidth - 10)/3,15,hWnd,(HMENU)3,hInst,NULL);
		SendMessage(MAIN_STATIC1,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_STATIC2,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_STATIC3,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_BUTTON1,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_PASSWORDEDIT,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_BUTTON2,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_RECORDSLIST,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_SIGNOPTION,WM_SETFONT,(WPARAM)MainFont,NULL);
		SendMessage(MAIN_LIST1,WM_SETFONT,(WPARAM)List1Font,NULL);
		SendMessage(MAIN_NAMECOMBO,WM_SETFONT,(WPARAM)List1Font,NULL);
		SendMessage(MAIN_PASSWORDEDIT,EM_SETPASSWORDCHAR,(WPARAM)42,NULL);
		COl.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
		COl.pszText = L"����";
		COl.cx = 100;
		//COl.iOrder = 0;
		ListView_InsertColumn(MAIN_RECORDSLIST, 0, &COl);
		COl.mask =  LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM | LVCF_SUBITEM;
		COl.pszText = L"ʱ��";
		COl.cx = 75;
		//COl.iOrder = 0;
		ListView_InsertColumn(MAIN_RECORDSLIST, 1, &COl);
		SystemParametersInfo(SPI_GETWORKAREA, 0, &rect, 0);
	    cx = rect.right - rect.left;
	    cy = rect.bottom - rect.top;
		//��ȡ�û�����
		//
		// 
		// ���棺��Ҫ���롣
		// 
		// 
		//  û�д���
		//��ȡǩ���ļ�
		if (hf == NULL) {
			if (CoInitialize(NULL) != S_OK) {
				MessageBox(hWnd, _T("��������޷�����Ĵ�����Ҫ�رգ��������������޷������û�������ɵ�"), _T("�쳣ָʾ"), 0);
				PostQuitMessage(-1);
			};
			XmlFile = XMLDOC.CreateInstance(__uuidof(MSXML2::DOMDocument60));
			if (!SUCCEEDED(XmlFile)) {
				MessageBox(hWnd, _T("��������޷�����Ĵ�����Ҫ�رգ��������������޷������û�������ɵ�"), _T("�쳣ָʾ"), 0);
				PostQuitMessage(-1);
			}
			GetEnvironmentVariable(_T("APPDATA"), szFile, 1024);
			wcscat_s(szFile, _T("\\Xupeng Studio\\Classsign\\Userprof.xupestd"));
			XMLDOC->load(szFile);
			XMLROOT = XMLDOC->GetdocumentElement();//��ø��ڵ�;
			XMLROOT->get_childNodes(&XMLNODES);//��ø��ڵ�������ӽڵ�;
			long XMLNODESNUM, ATTSNUM;
			XMLNODES->get_length(&XMLNODESNUM);//��������ӽڵ�ĸ���;
			for (int I = 0; I < XMLNODESNUM; I++)
			{
				XMLNODES->get_item(I, &XMLNODE);//���ĳ���ӽڵ�;
				XMLNODE->get_attributes(&XMLNODEATTS);//���ĳ���ڵ����������;
				XMLNODEATTS->get_length(&ATTSNUM);//����������Եĸ���;
				char b[2048] = {};
				for (int J = 0; J < ATTSNUM; J++)
				{
					XMLNODEATTS->get_item(J, &XMLNODE);//���ĳ������;
						
					if (XMLNODE->nodeName == (_bstr_t)"Username") {
						char a[2048] = {};
						TCHAR turned[1024] = L"";
						unsigned char key[16] = "madebyXupestd3.";
						char ques[16] = "";
						strcat(a, (char *)(_bstr_t)XMLNODE->text);
						int i = 0;
						for (i = 0; i <= 1024 / 16 && a[i * 16] != '\0'; i++)
						{
							//unsigned char x[12] = {};
							for (int j = 0; j < 4; j++)
							{
								for (int base64num = 0; base64num < 64; base64num++)
								{
									if (a[i * 16 + j * 4] == base64[base64num])ques[j * 3] = base64num * 4;
								}
								for (int base64num = 0; base64num < 64; base64num++)
								{
									if (a[i * 16 + j * 4 + 1] == base64[base64num])
									{
										ques[j * 3] = ques[j * 3] + base64num / 16;
										ques[j * 3 + 1] = base64num % 16 * 16;
									}
								}
								for (int base64num = 0; base64num < 64; base64num++)
								{
									if (a[i * 16 + j * 4 + 2] == base64[base64num])
									{
										ques[j * 3 + 1] = ques[j * 3 + 1] + base64num / 4;
										ques[j * 3 + 2] = base64num % 4 * 64;
									}

								}
								for (int base64num = 0; base64num < 64; base64num++)
								{
									if (a[i * 16 + j * 4 + 3] == base64[base64num])
									{
										ques[j * 3 + 2] = ques[j * 3 + 2] + base64num;
									}
								}
							}
							for (int k = 0; k <= 12; k++)b[i * 12 + k] = ques[k];
						}
						AES aes(key);
						aes.InvCipher(b, i * 16);
						MultiByteToWideChar(CP_UTF8, 0, b, sizeof(b), turned, sizeof(turned));
						ListBox_AddItemData(MAIN_LIST1,turned);
						ComboBox_AddItemData(MAIN_NAMECOMBO, turned);
						
					}
				}
			}
		}
		break;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		if (lParam == (LPARAM)MAIN_LIST1)
		{
			int Textlength = 0;
			int selectd = 0;
			int pos = 0;
			for (pos = 0; pos < ListBox_GetCount(MAIN_LIST1); pos++)
				if (ListBox_GetSel(MAIN_LIST1, pos))break;
			if (pos >= ListBox_GetCount(MAIN_LIST1))break;
			Textlength = SendMessage(MAIN_LIST1, LB_GETTEXTLEN, (WPARAM)pos, NULL);
			TCHAR* x = new TCHAR[Textlength+1];
			ZeroMemory(x, Textlength + 1);
			SendMessage(MAIN_LIST1, LB_GETTEXT, (WPARAM)pos, (LPARAM)x);
			SendMessage(MAIN_NAMECOMBO, CB_SETCURSEL, (WPARAM)pos, NULL);
			delete[] x;
		}
		else if (lParam == (LPARAM)MAIN_NAMECOMBO) {
			int Textlength = 0;
			int selectd = 0;
			int pos = 0;
			TCHAR* x;
			switch (wmEvent)
			{
				case CBN_SELCHANGE:
					pos = ComboBox_GetCurSel(MAIN_NAMECOMBO);
					SendMessage(MAIN_LIST1, LB_SETCURSEL, (WPARAM)pos, NULL);
					break;
				case CBN_EDITUPDATE:
					x = (TCHAR *)malloc((ComboBox_GetTextLength(MAIN_NAMECOMBO)+1)*sizeof(TCHAR));
					ComboBox_GetText(MAIN_NAMECOMBO, x, ComboBox_GetTextLength(MAIN_NAMECOMBO)+1);
					pos=ComboBox_FindStringExact(MAIN_NAMECOMBO, -1, x);
					if (pos >= ComboBox_GetCount(MAIN_NAMECOMBO)||pos<0) { free(x); break; }
					ComboBox_SetCurSel(MAIN_NAMECOMBO, pos);
					ListBox_SetCurSel(MAIN_LIST1, pos);
					free(x);
					break;
			}
		}
		else if(lParam==(LPARAM)MAIN_SIGNOPTION)
		{
			int selected = 0;
			
			time_t nowtime;
			struct tm t;
			char x[10]="";
			TCHAR y[10] = L"";
			TCHAR* result;
			time(&nowtime);
			localtime_s(&t, &nowtime);
			sprintf_s(x,10, "%s:%s:%s", pluszero(t.tm_hour), pluszero(t.tm_min), pluszero(t.tm_sec));
			MultiByteToWideChar(CP_ACP, NULL, x, 10, y, 10);
			int textLength = (ListBox_GetTextLen(MAIN_LIST1, selected) + 12);
			selected = ListBox_GetCurSel(MAIN_LIST1);
			if (selected == -1)break;
			result = (TCHAR*)malloc(textLength * sizeof(TCHAR));
			ZeroMemory(result, textLength);
			ListBox_GetText(MAIN_LIST1, selected, result);
			lvi.mask = LVIF_TEXT;
			lvi.iItem = ListView_GetItemCount(MAIN_RECORDSLIST);
			lvi.iSubItem = 0;
			lvi.pszText = result;
			ListView_InsertItem(MAIN_RECORDSLIST, &lvi);
			lvi.iSubItem = 1;
			lvi.pszText = y;
			ListView_SetItem(MAIN_RECORDSLIST, &lvi);
			//wcscat_s(result,textLength,L" \0");
			//wcscat_s(result, textLength, y);
			//ListBox_AddItemData(MAIN_RECORDSLIST, result);
			free(result);
			ListBox_DeleteString(MAIN_LIST1, selected);
			ComboBox_DeleteString(MAIN_NAMECOMBO, selected);
			SetWindowText(MAIN_NAMECOMBO, L"");
		}
		// �����˵�ѡ��:
		switch (wmId)
		{
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;
		case IDM_STARTQUICK:
			option = 1;
			DialogBox(hInst,MAKEINTRESOURCE(IDD_PASSWORDDIALOG),hWnd,Password);
			break;
		 case IDM_USER:
			DialogBox(hInst,MAKEINTRESOURCE(IDD_USERLOGGING),hWnd,Userlog);
			break;
		 case IDM_ADMINLOG:
			 option=2;
			 DialogBox(hInst,MAKEINTRESOURCE(IDD_PASSWORDDIALOG),hWnd,Password);
			 break;
		
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}

		break;
	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		
		//	// TODO: �ڴ���������ͼ����...
		EndPaint(hWnd, &ps);
		break;
	case WM_NOTIFY:
		/*wmId = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		if(wmEvent==LB)*/
		break;
	case WM_CTLCOLORSTATIC:
		 hdc = (HDC)wParam;
         SetTextColor(hdc, RGB(0, 0, 0));
         SetBkMode(hdc, TRANSPARENT);
         return (LRESULT)GetStockObject(NULL_BRUSH);
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	case WM_GETMINMAXINFO:
		MINMAXINFO *FormMinSize;
		FormMinSize= (MINMAXINFO *)lParam;
		xoff=423;yoff=400;
		FormMinSize->ptMinTrackSize.x=xoff;FormMinSize->ptMinTrackSize.y=yoff;
		break;
	case WM_SIZE:
		RECT Windowlocation;
		GetWindowRect(hWnd,&Windowlocation);
		if(wParam !=SIZE_MINIMIZED){
			formheight = HIWORD(lParam);
			formwidth = LOWORD(lParam);
			MoveWindow(MAIN_LIST1,0,25,(formwidth - 10)/3,formheight - 25,TRUE);
			MoveWindow(MAIN_RECORDSLIST,(formwidth - 10)/3*2+10,25,(formwidth-10)/3,formheight - 25,TRUE);
			MoveWindow(MAIN_STATIC1,0,0,(formwidth - 10)/3,25,TRUE);
			MoveWindow(MAIN_STATIC2,(formwidth - 10)/3+5,35,(formwidth - 10)/3,15,TRUE);
			MoveWindow(MAIN_BUTTON1,(formwidth - 10)/3+5,50,(formwidth - 10)/3,15,TRUE);
			MoveWindow(MAIN_BUTTON2,(formwidth - 10)/3+5,65,(formwidth - 10)/3,15,TRUE);
			MoveWindow(MAIN_NAMECOMBO,(formwidth - 10)/3+5,0,(formwidth - 10)/3,25,TRUE);
			MoveWindow(MAIN_SIGNOPTION,(formwidth - 10)/3+5,100,(formwidth - 10)/3,25,TRUE);
			MoveWindow(MAIN_PASSWORDEDIT,(formwidth - 10)/3+5,80,(formwidth - 10)/3,20,TRUE);
			MoveWindow(MAIN_STATIC3,(formwidth - 10)/3*2+10,0,(formwidth - 10)/3,15,TRUE);
		}
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

INT_PTR CALLBACK Start(HWND hDlg, UINT message,WPARAM wParam, LPARAM lParam){
	TCHAR szFile[4096];       // buffer for file name
	WIN32_FIND_DATA temp;
	HWND START_EDIT=GetDlgItem(hDlg,IDC_CHOOSEFILE);
	switch(message){
	case WM_COMMAND:
		switch(LOWORD(wParam))
		{
		case IDCANCEL:
			start=false;
			PostQuitMessage(0);	
			return (INT_PTR)TRUE;
			break;
		case IDOK:
			WCHAR Filename[1024];
			Filename[0]=L'\0';
			temp.cFileName[0]=L'\0';
			GetWindowText(START_EDIT, Filename,GetWindowTextLength(START_EDIT) + 1);
			if(Filename[0]==L'\0'){
				EndDialog(hDlg,(INT_PTR)TRUE);
				if (!InitInstance(hInst, 1))
				{
					return FALSE;
				}
				break;
			}
			FindFirstFile(Filename,&temp);
			if(temp.cFileName[0]==L'\0'){
				MessageBox(hDlg,_T("�Ҳ����ļ�����ȷ���ļ����Ƿ���ڣ�Ȼ����ȷ����"),_T("������ļ�������"),0);
				break;
			}
			start=true;
			hf = CreateFile(Filename,GENERIC_READ,FILE_SHARE_READ,NULL,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,NULL);
			
			EndDialog(hDlg,(INT_PTR)TRUE);
			if (!InitInstance(hInst, 1))
			{
				return FALSE;
			}
			break;
			//�ж��ļ��Ƿ���ڣ�����ʼǩ������Ҫ�޸ġ�
		case IDC_GETFILE:
			ZeroMemory(&classignmainfilename,sizeof(classignmainfilename));
			classignmainfilename.lStructSize=sizeof(classignmainfilename);
			classignmainfilename.hwndOwner=hDlg;
			classignmainfilename.lpstrFile=szFile;
			classignmainfilename.lpstrFile[0] = '\0';
			classignmainfilename.nMaxFile = sizeof(szFile);
			classignmainfilename.lpstrFilter = _T("�����ļ������ܲ����ݣ�\0*.*\0�༶ǩ����ר���ļ�\0*.TXT\0");
			classignmainfilename.nFilterIndex=1;
			classignmainfilename.lpstrFileTitle=NULL;
			classignmainfilename.nMaxFileTitle=0;
			classignmainfilename.lpstrInitialDir = NULL;
			classignmainfilename.Flags=OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
			if (GetOpenFileName(&classignmainfilename)==TRUE) 
			{
				SetWindowText(START_EDIT,classignmainfilename.lpstrFile);
			}
		}
		
	}
	return (INT_PTR)FALSE;
}



//����Ա�������������......
INT_PTR CALLBACK Password(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:
			EndDialog(hDlg,(INT_PTR)TRUE);
			if (option==2){
				DialogBox(hInst,MAKEINTRESOURCE(IDD_ADMIN),NULL,Admin);
			}
			break;
		case IDCANCEL:
			EndDialog(hDlg,(INT_PTR)TRUE);
			break;
		}
		break;
	}
	return (INT_PTR)FALSE;
}


//�û�������棬��Ҫ�ȴ����
INT_PTR CALLBACK UserProcess(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}




//�û���¼���棬�������......
INT_PTR CALLBACK Userlog(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		switch(LOWORD(wParam))
		{
		case IDCANCEL:
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		case IDOK:
			EndDialog(hDlg,LOWORD(wParam));
			DialogBox(hInst,MAKEINTRESOURCE(IDD_USERINFORMATION),NULL,UserProcess);
			return (INT_PTR)TRUE;
		case IDC_FORGET:
			MessageBox(hDlg,_T("����ϵ����Ա�����������롣"),_T("�������룿"),0);
			break;
		}
		break;
	}
	return (INT_PTR)FALSE;
}



//����Ա����ҳ�棬��Ҫ����д����
INT_PTR CALLBACK Admin(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	HWND ADMIN_TABCONTROL;
	ADMIN_TABCONTROL= GetDlgItem(hDlg, IDC_MAINTAB);
	wchar_t pItem[256] = { 0 };
	LPWSTR tabname[5] = { L"ѧ������",L"ǩ���ļ�����",L"ǩ���ƻ�����",L"�߼�����"}; //����һ����ά���� ���tab��ǩ����L"����",L"ʱ���"
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		TCITEM tie;//����tab��ǩ������
		//���忪ʼ���� tie���ֶ� Mask psztext,ccxtextmax,image,lparam
		tie.mask = TCIF_TEXT;//psztext�ֶ���Ч
		for (INT i = 0; i < 4; i++)
		{
			tie.pszText = tabname[i];
			TabCtrl_InsertItem(ADMIN_TABCONTROL, i, &tie);
		}
		RECT rect;//���tab�ؼ�������λ��
		GetClientRect(ADMIN_TABCONTROL, &rect);
		// ������������ tab�ؼ�λ���ƶ�
		AdminTab[0] = CreateDialog(hInst, MAKEINTRESOURCE(IDD_ADMIN_TAB_STUDENT), ADMIN_TABCONTROL, Student);
		MoveWindow(AdminTab[0], 2, 29, rect.right - rect.left - 6, rect.bottom - rect.top - 35, FALSE);
		AdminTab[1] = CreateDialog(hInst, MAKEINTRESOURCE(IDD_ADMIN_TAB_SIGNFILE), ADMIN_TABCONTROL, SignFile);
		MoveWindow(AdminTab[1], 2, 29, rect.right - rect.left - 6, rect.bottom - rect.top - 35, FALSE);
		AdminTab[2] = CreateDialog(hInst, MAKEINTRESOURCE(IDD_ADMIN_TAB_BOOKING), ADMIN_TABCONTROL, Booking);
		MoveWindow(AdminTab[2], 2, 29, rect.right - rect.left - 6, rect.bottom - rect.top - 35, FALSE);
		AdminTab[3] = CreateDialog(hInst, MAKEINTRESOURCE(IDD_ADMIN_TAB_ADVANCE), ADMIN_TABCONTROL, Advance);
		MoveWindow(AdminTab[3], 2, 29, rect.right - rect.left - 6, rect.bottom - rect.top - 35, FALSE);
		/* δ�����ӵ����ݡ�
		hDlg_intab[4]=CreateDialog(hInst,MAKEINTRESOURCE(IDD_TIMETABLE),htabctrl,PageE);
		MoveWindow(hDlg_intab[4],2,29,rect.right - rect.left-6,rect.bottom - rect.top-35,FALSE);
		*/
		ShowWindow(AdminTab[0], TRUE);
		SetForegroundWindow(hDlg);

		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	case WM_NOTIFY:		 //TAB�ؼ��л�����ʱ���͵���Ϣ
		if ((INT)wParam == IDC_MAINTAB) //����Ҳ������һ��NMHDR *nm = (NMHDR *)lParam���ָ������ȡ ������¼�
		{					//���߿����в���NMHDR�ṹ
			if (((LPNMHDR)lParam)->code == TCN_SELCHANGE) //��TAB��ǩת����ʱ����TCN_SELCHANGE��Ϣ
			{

				int sel = TabCtrl_GetCurSel(ADMIN_TABCONTROL);
				ShowWindow(AdminTab[0], sel == 0); //��ʾ������ShowWindow����
				ShowWindow(AdminTab[1], sel == 1);
				ShowWindow(AdminTab[2], sel == 2);
				ShowWindow(AdminTab[3], sel == 3);
				/*
				ShowWindow(hDlg_intab[4],sel==4);*/
			}
		}
		break;
	

	
	}
	return (INT_PTR)FALSE;
}


INT_PTR CALLBACK Advance(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

INT_PTR CALLBACK Booking(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

INT_PTR CALLBACK Student(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}

INT_PTR CALLBACK SignFile(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}


HANDLE tempfirstfile;
//��ʼ������һ��
INT_PTR CALLBACK SetPassword(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	TCHAR szBuffer[2048];
	DWORD dw;
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		
		GetEnvironmentVariable(_T("APPDATA"), szBuffer, 2048);
		wcscat_s(szBuffer, L"\\Xupeng Studio");
		CreateDirectory(szBuffer,NULL);
		wcscat_s(szBuffer, L"\\Classsign");
		CreateDirectory(szBuffer, NULL);
		wcscat_s(szBuffer, _T("\\Userprof.xupestd"));
		tempfirstfile = CreateFile(szBuffer,GENERIC_ALL, 0, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
		CloseHandle(tempfirstfile);
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		switch (LOWORD(wParam)) {
		case IDOK:
		{
			HWND STEP1_PSW = GetDlgItem(hDlg, STEP1_FIRSTPSW);
			HWND STEP1_PSW2 = GetDlgItem(hDlg, STEP1_FIRSTPSW2);
			TCHAR cmp1[2048] = { '\0' }, cmp2[2048] = { '\0' };
			GetWindowText(STEP1_PSW, cmp1, GetWindowTextLength(STEP1_PSW)+1);
			GetWindowText(STEP1_PSW2, cmp2, GetWindowTextLength(STEP1_PSW2)+1);
			if (wcscmp(cmp1, cmp2) != 0) {
				MessageBox(hDlg, L"������������Ĳ�һ�£����������롣", L"���벻һ�´���", 0);
				SetWindowText(STEP1_PSW, _T(""));
				SetWindowText(STEP1_PSW2, _T(""));
				break;
			}
			else {
				if (cmp1[0] == '\0') {
					MessageBox(hDlg, L"����Ϊ�գ����������롣", L"���������", 0);
					break;
				}
				//AES�����㷨
				unsigned char key[16] = {};
				unsigned char encode[16] = "MadeByxupengstd";
				for (int total = 0; total <= (int)(wcslen(cmp1)/16); total++) {
					key[0] = '\0';
					int num = 0;
					for (num = 0; num < 16; num++) {
						if (cmp1[total * 16 + num])
						{
							key[num] = cmp1[total * 16 + num];
						}
						else {
							break;
						}
					}
					for (;num<16;num++){
						key[num] = 0;
					}
					AES aes(key);
					aes.Cipher(encode);
				}
				MSXML2::IXMLDOMDocumentPtr pXMLDOC;
				MSXML2::IXMLDOMElementPtr pXMLroot;
				MSXML2::IXMLDOMElementPtr pXMLsettings;
				if (CoInitialize(NULL)!=S_OK ){
					MessageBox(hDlg, _T("��������޷�����Ĵ�����Ҫ�رգ��������������޷������û�������ɵ�"), _T("�쳣ָʾ"), 0);
					PostQuitMessage(-1);
				};
				HRESULT XmlFile = pXMLDOC.CreateInstance(__uuidof(MSXML2::DOMDocument60));
				if (!SUCCEEDED(XmlFile)) {
					MessageBox(hDlg, _T("��������޷�����Ĵ�����Ҫ�رգ��������������޷������û�������ɵ�"), _T("�쳣ָʾ"), 0);
					PostQuitMessage(-1);
				}
				TCHAR baseend[19] = L"";
				for (int i = 0; i < 5; i++) {
					baseend[i * 4 + 0] = base64[(unsigned int)(encode[i * 3] / 4)];
					baseend[i * 4 + 1] = base64[(unsigned int)(encode[i * 3] % 4) * 16 + (unsigned int)(encode[i * 3 + 1] / 16)];
					baseend[i * 4 + 2] = base64[(unsigned int)(encode[i * 3 + 1] % 16) * 4 + (unsigned int)(encode[i * 3 + 2] / 64)];
					baseend[i * 4 + 3] = base64[(unsigned int)(encode[i * 3 + 2] % 64)];
				}
				baseend[16] = base64[(unsigned int)encode[15] / 4];
				baseend[17] = base64[(unsigned int)encode[15] % 4 * 16];
				baseend[18] = 0;
				TCHAR* answer;
				answer = baseend;
				GetEnvironmentVariable(_T("APPDATA"), szBuffer, 2048);
				wcscat_s(szBuffer, _T("\\Xupeng Studio\\Classsign\\Userprof.xupestd"));
				
				pXMLDOC->raw_createElement((bstr_t)(char*)("UserData"), &pXMLroot);
				pXMLDOC->raw_createElement((bstr_t)(char*)("password"), &pXMLsettings);
				pXMLDOC->raw_appendChild(pXMLroot, NULL);
				
				pXMLroot->raw_appendChild(pXMLsettings, NULL);
				
				pXMLsettings->Puttext((bstr_t)answer);
				pXMLDOC->save(szBuffer);

				CoUninitialize();
				/*AES�����㷨
				for (int total = (int)(wcslen(cmp1)/16); total >= 0 ; total--) {
					key[0] = '\0';
					int num = 0;
					for (num = 0; num < 16; num++) {
						if (cmp1[total * 16 + num])
						{
							key[num] = cmp1[total * 16 + num];
						}
						else {
							break;
						}
					}
					for (; num < 16; num++) {
						key[num] = 0;
					}
					AES aes(key);
					aes.InvCipher(encode);
				}
				ZeroMemory(cmp1,2048);
				ZeroMemory(cmp2, 2048);
				*/
				EndDialog(hDlg, LOWORD(wParam));
				DialogBox(hInst, MAKEINTRESOURCE(FIRSTRUN_STEP2), NULL, SetNames);

			}
			return (INT_PTR)TRUE;
		}
		break;
		case IDCANCEL:
			GetEnvironmentVariable(_T("APPDATA"), szBuffer, 2048);
			wcscat_s(szBuffer, L"\\Xupeng Studio\\classsign");
			wcscat_s(szBuffer, _T("\\Userprof.xupestd"));
			
			DeleteFile(szBuffer);
			dw = GetLastError();
			start = false;
			PostQuitMessage(0);
			break;
		}

		break;

	}
	
	return (INT_PTR)FALSE;
}



//��ʼ�����ڶ�����
INT_PTR CALLBACK SetNames(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	ofstream ofs;
	UNREFERENCED_PARAMETER(lParam);
	HWND FIRSTRUN_MAINLIST1 = GetDlgItem(hDlg, FIRSTRUN_MAINLIST);
	int FirstrunMainListMenuSelected = 0;
	HMENU FirstrunMainlistMenu = CreatePopupMenu();
	LVITEM lvi;
	TCHAR szFile[1024];
	int itemCount = 0, i = 0;
	DWORD number = 0,FileReadSize;
	BYTE filetemp[32] = {0};
	HANDLE ImportFile;
	vector <unsigned char> Username;
	DWORD FileSize,readedFileSize=0;
	int TextEncodeMode; HRESULT XmlFile;
	int textlength;
	TCHAR temp[1024]=_T("\0");
	CHAR templeChar[2048] = "\0";
	vector <TCHAR> usernametemps;
	bool filebomisjudged=false;
	MSXML2::IXMLDOMDocumentPtr pXMLDOC = NULL;
	MSXML2::IXMLDOMElementPtr pXMLroot;
	MSXML2::IXMLDOMElementPtr pXMLsettings;
	switch (message)
	{
	case WM_INITDIALOG:
		LV_COLUMN lvc;
		lvc.mask = LVCF_TEXT;
		lvc.pszText = (LPTSTR)L"����";
		lvc.cxMin = 100;

		lvc.cx = 100;
		ListView_InsertColumn(FIRSTRUN_MAINLIST1, 0, &lvc);
		
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if(HIWORD(wParam)==BN_CLICKED){
			switch (LOWORD(wParam)) {
			case IDOK:
				if (CoInitialize(NULL) != S_OK) {
					MessageBox(hDlg, _T("��������޷�����Ĵ�����Ҫ�رգ��������������޷������û�������ɵ�"), _T("�쳣ָʾ"), 0);
					PostQuitMessage(-1);
				};
				XmlFile = pXMLDOC.CreateInstance(__uuidof(MSXML2::DOMDocument60));
				if (!SUCCEEDED(XmlFile)) {
					MessageBox(hDlg, _T("��������޷�����Ĵ�����Ҫ�رգ��������������޷������û�������ɵ�"), _T("�쳣ָʾ"), 0);
					PostQuitMessage(-1);
				}
				GetEnvironmentVariable(_T("APPDATA"), szFile, 1024);
				wcscat_s(szFile, _T("\\Xupeng Studio\\Classsign\\Userprof.xupestd"));
				itemCount = ListView_GetItemCount(FIRSTRUN_MAINLIST1);
				pXMLDOC->load((variant_t)szFile);
				pXMLroot = pXMLDOC->documentElement;
				for (i = 0; i < itemCount; i++)
				{
					ZeroMemory(templeChar, sizeof(templeChar));
					pXMLDOC->raw_createElement((_bstr_t)(char*)"UserName", &pXMLsettings);
					lvi.mask = LVIF_TEXT;
					lvi.iItem = i;
					lvi.iSubItem = 0;
					lvi.pszText = temp;
					ListView_GetItemText(FIRSTRUN_MAINLIST1, i, 0, temp, 1024);
					//int x=SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMTEXT, (WPARAM)i, (LPARAM)&lvi);
					//x = GetLastError();
					unsigned char key[16] = "madebyXupestd3.";
					AES aes(key);
					
					int number = WideCharToMultiByte(CP_UTF8, 0, temp, -1, templeChar, 0, NULL, NULL);
					WideCharToMultiByte(CP_UTF8, 0, temp, -1, templeChar, number + 1, NULL, NULL);
					int strlength = strlen(templeChar);
					aes.Cipher(templeChar);
					CHAR output[2048]="";
					for (int j = 0; j < (strlength / 12 + 1)*16/12+1; j++)
					{
						for (int k = 0; k < 4; k++) {
							output[j * 16 + k * 4] = base64[((unsigned char)templeChar[j * 12 + k * 3] >> 2) % 64];
							output[j * 16 + k * 4 + 1] = base64[(((unsigned char)templeChar[j * 12 + k * 3] << 4) + ((unsigned char)templeChar[j * 12 + k * 3 + 1] >> 4)) % 64];
							output[j * 16 + k * 4 + 2] = base64[(((unsigned char)templeChar[j * 12 + k * 3 + 1] << 2) + ((unsigned char)templeChar[j * 12+k * 3 + 2] >> 6)) % 64];
							output[j * 16 + k * 4 + 3] = base64[((unsigned char)templeChar[j * 12 + k * 3 + 2]) % 64];
						}
					}
					pXMLsettings->setAttribute((_bstr_t)L"Username",output);
					pXMLsettings->setAttribute((_bstr_t)L"Account", 10000 + i);
					pXMLroot->raw_appendChild(pXMLsettings, NULL);
				}
				pXMLDOC->save(szFile);
				/*
				pXMLroot->raw_appendChild(pXMLsettings, NULL);

				pXMLsettings->Puttext((bstr_t)answer);
				pXMLDOC->save(szBuffer);*/
				CoUninitialize();
				start = true;
				SHGetFolderPath(hDlg, CSIDL_DESKTOP, 0, 0, szFile);
				wcscat_s(szFile, _T("\\�༶ǩ�����û��˺��ļ�.csv"));
				ofs.open(szFile);
				itemCount = ListView_GetItemCount(FIRSTRUN_MAINLIST1);
				ofs << "�û�ID,�û���,����\n";
				for (i = 0; i < itemCount; i++)
				{
					ZeroMemory(temp,sizeof(temp));
					ZeroMemory(templeChar, sizeof(templeChar));
					ListView_GetItemText(FIRSTRUN_MAINLIST1, i, 0, temp, 1024);
					WideCharToMultiByte(CP_ACP, NULL, temp, 1024, templeChar, 2048, NULL, NULL);
					ofs << 10000 + i << ",";
					for(int k =0;k<2048;k++)
						if(*(templeChar + k))
							ofs << *(templeChar+k);
					ofs << ",123456\n";
				}
				EndDialog(hDlg, LOWORD(wParam));
				DialogBox(hInst, MAKEINTRESOURCE(IDD_STARTDIALOG), NULL, Start);
				
				return (INT_PTR)TRUE;
			case FIRSTRUN_NEW:
				itemCount = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
				lvi.mask = LVIF_TEXT;
				lvi.iItem = itemCount;
				lvi.iSubItem = 0;
				lvi.pszText = L"������";
				SetFocus(FIRSTRUN_MAINLIST1);
				ListView_InsertItem(FIRSTRUN_MAINLIST1,&lvi);
				SendMessage(FIRSTRUN_MAINLIST1, LVM_EDITLABEL, (WPARAM)lvi.iItem, NULL);
				return (INT_PTR)TRUE;
			case FIRSTRUN_DELETEALL:
				ListView_DeleteAllItems(FIRSTRUN_MAINLIST1);
				return (INT_PTR)TRUE;
			case FIRSTRUN_IMPORT:

				ZeroMemory(&classignmainfilename, sizeof(classignmainfilename));
				classignmainfilename.lStructSize = sizeof(classignmainfilename);
				classignmainfilename.hwndOwner = hDlg;
				classignmainfilename.lpstrFile = szFile;
				classignmainfilename.lpstrFile[0] = '\0';
				classignmainfilename.nMaxFile = sizeof(szFile);
				classignmainfilename.lpstrFilter = _T("�ı��ļ�\0*.txt\0Excel ���\0*.xlsx\0");
				classignmainfilename.nFilterIndex = 1;
				classignmainfilename.lpstrFileTitle = NULL;
				classignmainfilename.nMaxFileTitle = 0;
				classignmainfilename.lpstrInitialDir = NULL;
				classignmainfilename.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
				if (GetOpenFileName(&classignmainfilename) == TRUE)
				{
					switch (classignmainfilename.nFilterIndex)
					{
					case 1://����txt�ĵ�ʱ�Ĳ�����

						ImportFile = CreateFile(classignmainfilename.lpstrFile, GENERIC_ALL, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
						if (ImportFile == INVALID_HANDLE_VALUE)
						{
							MessageBox(hDlg, L"�ļ��޷��򿪣�����ԭ��.", L"�ļ���ʧ��", 0);
							return (INT_PTR)FALSE;
						}

						FileSize = GetFileSize(ImportFile,NULL);
						
						while (true) {
							i = 0; int pos = 0;
							int readSize = 32;

							readSize = FileSize - readedFileSize > 32 ? 32 : FileSize - readedFileSize;
							if (FALSE == ReadFile(ImportFile, filetemp, readSize, &FileReadSize, NULL)) {
								MessageBox(hDlg, L"û���ļ���ȡȨ�ޣ����鲢�ų�����", L"�ļ���ȡʧ��", 0);
								return (INT_PTR)FALSE;
							}
							while (pos < readSize) {
								if (filetemp[pos] != '\r' && filetemp[pos] != '\n') {
									Username.push_back(filetemp[pos]);
								}
								else {
									if (Username.size() == 0) {
										pos++;
										continue;
									}
									else {
										//�����ж��ļ��ı����ʽ
										TextEncodeMode = IsTextUnicode(Username.data(), 32, 0);
										
										//�ж��ļ��Ƿ���UTF-8����
										if (!TextEncodeMode) {
											unsigned int judgepos = 0;
											while (judgepos < Username.size()) {
												if (Username[judgepos] < 0x7F)
												{
													judgepos++;
												}
												else if (Username[judgepos] >= (0xC0) && Username[judgepos] < (0xE0))
												{
													if (Username[judgepos + 1] >= (0x80) && Username[judgepos + 1] < (0xC0)) {
														judgepos += 2;
													}
													else {
														break;
													}
												}
												else if (Username[judgepos] >= (0xE0) && Username[judgepos] < (0xF0)) {
													if (Username[judgepos + 1] >= (0x80) && Username[judgepos + 1] < (0xC0)) {
														if (Username[judgepos + 2] >= (0x80) && Username[judgepos + 2] < (0xC0)) {
															judgepos = judgepos + 3;
														}
														else {
															break;
														}
													}
													else {
														break;
													}
												}
												else
												{
													break;
												}
											}
											if (judgepos >= Username.size())TextEncodeMode = 4;
										}
										//�ж��ַ���������GBK
										if (!TextEncodeMode) {

											unsigned int judgepos = 0;
											while (judgepos < Username.size()) {
												if (Username[judgepos] > 0x8E && Username[judgepos] <= 0xFF) {
													if (Username[judgepos + 1] >= 0x40 && Username[judgepos + 1] < 0xFF) {
														judgepos = judgepos + 2;
														continue;
													}
													else {
														TextEncodeMode = 0;
														break;
													}
												}
												else if (Username[judgepos] < 0x80)
												{
													judgepos++;
												}
												else { break; }
											}
											if (GetACP() == 936 && (Username.size()<=judgepos)) TextEncodeMode = 2;
										}
										int length = Username.size();
										CHAR* charnametemps = new CHAR[length+1];
										TCHAR* temple;
										switch (TextEncodeMode) {
										case 1://��ȡUnicode����Ĵ���
											for (unsigned int Usernumberjump = 0; Usernumberjump < Username.size(); Usernumberjump++)
											{
												usernametemps.push_back(Username[Usernumberjump]);
												if (Username[Usernumberjump] / 128 % 2 != 0)
												{
													usernametemps[Usernumberjump] = usernametemps[Usernumberjump] * 256 + usernametemps[Usernumberjump + 1];
													Usernumberjump++;
												}
											}
											usernametemps.push_back('\0');
											lvi.iItem = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
											lvi.iSubItem = 0;
											lvi.mask = LVIF_TEXT;
											lvi.pszText = usernametemps.data();
											ListView_InsertItem(FIRSTRUN_MAINLIST1, &lvi);
											break;
										case 2://��ȡGBK�Ĵ���
											Username.push_back('\0');
											for (int position = 0; position < length; position++) *(charnametemps + position) = Username[position];
											*(charnametemps + length) = '\0';
											length = MultiByteToWideChar(CP_ACP, 0, charnametemps, -1, NULL, 0);
											temple = new TCHAR[length];
											*(temple + length - 1) = '\0';
											length = MultiByteToWideChar(CP_ACP, 0, charnametemps, strlen(charnametemps), temple, wcslen(temple));  //GBK����ת����ASCII����
											lvi.mask = LVIF_TEXT;
											lvi.iItem = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
											lvi.iSubItem = 0;
											lvi.pszText = temple;

											ListView_InsertItem(FIRSTRUN_MAINLIST1, &lvi);
											delete[] temple;
											break;
										case 4://��ȡUTF-8�Ĵ���
											Username.push_back('\0');
											for (int position = 0; position < length; position++) *(charnametemps + position) = Username[position];
											*(charnametemps + length) = '\0';
											length = MultiByteToWideChar(CP_UTF8, 0, charnametemps, -1, NULL, 0);
											temple = new TCHAR[length];
											*(temple + length - 1) = '\0';
											length = MultiByteToWideChar(CP_UTF8, 0, charnametemps, strlen(charnametemps), temple, wcslen(temple));  //GBK����ת����ASCII����
											lvi.mask = LVIF_TEXT;
											lvi.iItem = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
											lvi.iSubItem = 0;
											lvi.pszText = temple;
											ListView_InsertItem(FIRSTRUN_MAINLIST1, &lvi);
											delete[] temple;
											break;
										}
										Username.clear();
										usernametemps.clear();
									}
								}
								pos++;
							}
							readedFileSize += FileReadSize;
							if (readedFileSize >= FileSize)
							{
								//�����ж��ļ��ı����ʽ
								TextEncodeMode = IsTextUnicode(Username.data(), 32, 0);

								//�ж��ļ��Ƿ���UTF-8����
								if (!TextEncodeMode) {
									unsigned int judgepos = 0;
									while (judgepos < Username.size()) {
										if (Username[judgepos] < (0x7F))
										{
											judgepos++;
										}
										else if (Username[judgepos] >= (0xC0) && Username[judgepos] < (0xE0))
										{
											if (Username[judgepos + 1] >= (0x80) && Username[judgepos + 1] < (0xC0)) {
												judgepos += 2;
											}
											else {
												break;
											}
										}
										else if (Username[judgepos] >= (0xE0) && Username[judgepos] < (0xF0)) {
											if (Username[judgepos + 1] >= (0x80) && Username[judgepos + 1] < (0xC0)) {
												if (Username[judgepos + 2] >= (0x80) && Username[judgepos + 2] < (0xC0)) {
													judgepos = judgepos + 3;
												}
												else {
													break;
												}
											}
											else {
												break;
											}
										}
										else
										{
											break;
										}
									}
									if (judgepos >= Username.size())TextEncodeMode = 4;
								}
								//�ж��ַ���������GBK
								if (!TextEncodeMode) {

									unsigned int judgepos = 0;
									while (judgepos < Username.size()) {
										if (Username[judgepos] > 0x8E && Username[judgepos] <= 0xFF) {
											if (Username[judgepos + 1] >= 0x40 && Username[judgepos + 1] < 0xFF) {
												judgepos = judgepos + 2;
												continue;
											}
											else {
												TextEncodeMode = 0;
												break;
											}
										}
										else if (Username[judgepos] < 0x80)
										{
											judgepos++;
										}
										else { break; }
									}
									if (GetACP() == 936 && judgepos >= Username.size()) TextEncodeMode = 2;
								}
								int length = Username.size();
								char* charnametemps = new char[length+1];
								TCHAR* temple;
								switch (TextEncodeMode) {
								case 1://��ȡUnicode����Ĵ���
									for (unsigned int Usernumberjump = 0; Usernumberjump < Username.size(); Usernumberjump++)
									{
										usernametemps.push_back(Username[Usernumberjump]);
										if (Username[Usernumberjump] / 128 % 2 != 0)
										{
											usernametemps[Usernumberjump] = usernametemps[Usernumberjump] * 256 + usernametemps[Usernumberjump + 1];
											Usernumberjump++;
										}
									}
									usernametemps.push_back('\0');
									lvi.iItem = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
									lvi.iSubItem = 0;
									lvi.mask = LVIF_TEXT;
									lvi.pszText = usernametemps.data();
									ListView_InsertItem(FIRSTRUN_MAINLIST1, &lvi);
									break;
								case 2://��ȡGBK�Ĵ���
									
									//Username.push_back('\0');
									
									for (int position = 0; position < length; position++) *(charnametemps + position) = Username[position];
									*(charnametemps + length) = '\0';
									length = MultiByteToWideChar(CP_ACP, 0, charnametemps, -1, NULL, 0);
									temple = new TCHAR[length];
									*(temple + length - 1) = '\0';
									length = MultiByteToWideChar(CP_ACP, 0, charnametemps, strlen(charnametemps), temple, wcslen(temple));  //GBK����ת����ASCII����
									lvi.mask = LVIF_TEXT;
									lvi.iItem = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
									lvi.iSubItem = 0;
									lvi.pszText = temple;
									ListView_InsertItem(FIRSTRUN_MAINLIST1, &lvi);
									delete[] temple;
									break;
								case 4://��ȡUTF-8�Ĵ���
									Username.push_back('\0');
									for (int position = 0; position < length; position++) *(charnametemps + position) = Username[position];
									length = MultiByteToWideChar(CP_UTF8, 0, charnametemps, -1, NULL, 0);
									temple = new TCHAR[length];
									*(temple + length - 1) = '\0';
									length = MultiByteToWideChar(CP_UTF8, 0, charnametemps, strlen(charnametemps), temple, wcslen(temple));  //GBK����ת����ASCII����
									lvi.mask = LVIF_TEXT;
									lvi.iItem = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
									lvi.iSubItem = 0;
									lvi.pszText = temple;
									ListView_InsertItem(FIRSTRUN_MAINLIST1, &lvi);
									delete[] temple;
									break;
								}
								break;
							}
						}
					
					CloseHandle(ImportFile);
					break;
					case 2:
						MessageBox(hDlg, _T("Beta 1���԰棬�����ڴ��������ݲ�֧�֡�"), _T("�����ڴ�"), 0);
						break;
					}
				}//������txt�ļ���ȡ����
				return (INT_PTR)TRUE;
			case IDCANCEL:
				GetEnvironmentVariable(_T("APPDATA"), szFile, 2048);
				wcscat_s(szFile, L"\\Xupeng Studio\\classsign");
				wcscat_s(szFile, _T("\\Userprof.xupestd"));

				DeleteFile(szFile);
				start = false;
				PostQuitMessage(0);
				EndDialog(hDlg, LOWORD(wParam));
				return (INT_PTR)TRUE;
			}
		}
	case WM_CONTEXTMENU:
		if((HWND)wParam==FIRSTRUN_MAINLIST1){
			AppendMenu(FirstrunMainlistMenu, MF_STRING,10000, L"�½�");
			FirstrunMainListMenuSelected = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETSELECTEDCOUNT, NULL, NULL);
			if (FirstrunMainListMenuSelected!=0) {
				AppendMenu(FirstrunMainlistMenu, MF_STRING, 10001, L"ɾ��");
				if (FirstrunMainListMenuSelected == 1)AppendMenu(FirstrunMainlistMenu, MF_STRING, 10002, L"������");
			}
			switch (TrackPopupMenu(FirstrunMainlistMenu, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), 0, FIRSTRUN_MAINLIST1, NULL))
			{
			case 10000:
				lvi.mask = LVIF_TEXT;
				lvi.iItem = SendMessage(FIRSTRUN_MAINLIST1, LVM_GETITEMCOUNT, NULL, NULL);
				lvi.iSubItem = 0;
				lvi.pszText = L"������";
				SetFocus(FIRSTRUN_MAINLIST1);
				ListView_InsertItem(FIRSTRUN_MAINLIST1, &lvi);
				SendMessage(FIRSTRUN_MAINLIST1, LVM_EDITLABEL, (WPARAM)lvi.iItem, NULL);
				return (INT_PTR)TRUE;
			case 10001:
				itemCount = ListView_GetItemCount(FIRSTRUN_MAINLIST1);
				for (i = 0; i < itemCount; i++) {
					if (ListView_GetItemState(FIRSTRUN_MAINLIST1, i, LVIS_SELECTED))
					{
						ListView_DeleteItem(FIRSTRUN_MAINLIST1, i);
						i--; itemCount--;
					}
				}
				return (INT_PTR)TRUE;
			case 10002:
				itemCount = ListView_GetItemCount(FIRSTRUN_MAINLIST1);
				for (i = 0; i < itemCount; i++) {
					if (ListView_GetItemState(FIRSTRUN_MAINLIST1, i, LVIS_SELECTED))break;
				}
				ListView_EditLabel(FIRSTRUN_MAINLIST1, i);
				return (INT_PTR)TRUE;
			}
		}
		break;

	case WM_NOTIFY:
		switch (((LPNMHDR)lParam)->code) {
		case LVN_ENDLABELEDIT:
			switch (((LPNMHDR)lParam)->idFrom) {
			case FIRSTRUN_MAINLIST:
				NMLVDISPINFO dpi;
				dpi = *((LPNMLVDISPINFOW)lParam);
				ListView_SetItem(FIRSTRUN_MAINLIST1, &dpi.item);
				break;
			}
			break;
		}
	}
	return (INT_PTR)FALSE;
}





//���ڶԻ���Ĳ���������ɡ�

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
CHAR* pluszero(int num)
{
	char *result;
	result = (char*)malloc(sizeof(char) * 3);
	ZeroMemory(result, sizeof(char) * 3);
	if (num < 10)
		sprintf(result, "0%d", num);
	else
		sprintf(result, "%d", num);
	return result;
	free(result);

	
}
