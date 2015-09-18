#include <windows.h>
#include <gdiplus.h>
#include <tchar.h>  
#include <stdio.h>
#include <Unknwn.h>
#include <Richedit.h>
#include <RichOle.h>
#include <assert.h>
#include "resource.h"

//#include <string>
#define IDC_IMAGE	202
#define ID_EDITCHILD	203
#define X_COORD	300
#define Y_COORD	200
#define ID_BUTTON_CLOSE	204
#define ID_BUTTON_MIN		205
#define ID_BUTTON_INSTALL	206
#define ID_BUTTON_CANCEL	207
#define ID_MY_BITMAP	2000
#define ID_DELL_BITMAP	2001

using namespace Gdiplus;

const char g_szClassName[] = "myWindowClass";

class CImageDataObject : IDataObject
{
public:
	// This static function accepts a pointer to IRochEditOle
	// and the bitmap handle.
	// After that the function insert the image in the current
	// position of the RichEdit
	//
	static void InsertBitmap(IRichEditOle* pRichEditOle,
		HBITMAP hBitmap);
private:
	ULONG m_ulRefCnt;
	BOOL m_bRelease;
	// The data being bassed to the richedit
	//
	STGMEDIUM m_stgmed;
	FORMATETC m_fromat;
public:
	CImageDataObject() : m_ulRefCnt(0) {
		m_bRelease = FALSE;
	}
	~CImageDataObject() {
		if (m_bRelease)
			::ReleaseStgMedium(&m_stgmed);
	}
	// Methods of the IUnknown interface
	//
	STDMETHOD(QueryInterface)(REFIID iid, void ** ppvObject)
	{
		if (iid == IID_IUnknown || iid == IID_IDataObject)
		{
			*ppvObject = this;
			AddRef();
			return S_OK;
		}
		else
			return E_NOINTERFACE;
	}
	STDMETHOD_(ULONG, AddRef)(void)
	{
		m_ulRefCnt++;
		return m_ulRefCnt;
	}
	STDMETHOD_(ULONG, Release)(void)
	{
		if (--m_ulRefCnt == 0)
		{
			delete this;
		}
		return m_ulRefCnt;
	}
	// Methods of the IDataObject Interface
	//
	STDMETHOD(GetData)(FORMATETC *pformatetcIn,
		STGMEDIUM *pmedium) {
			HANDLE hDst;
			hDst = ::OleDuplicateData(m_stgmed.hBitmap,
				CF_BITMAP, NULL);
			if (hDst == NULL)
			{
				return E_HANDLE;
			}
			pmedium->tymed = TYMED_GDI;
			pmedium->hBitmap = (HBITMAP)hDst;
			pmedium->pUnkForRelease = NULL;
			return S_OK;
	}
	STDMETHOD(GetDataHere)(FORMATETC* pformatetc,
		STGMEDIUM* pmedium ) {
			return E_NOTIMPL;
	}
	STDMETHOD(QueryGetData)(FORMATETC* pformatetc ) {
		return E_NOTIMPL;
	}
	STDMETHOD(GetCanonicalFormatEtc)(FORMATETC* pformatectIn ,
		FORMATETC* pformatetcOut ) {
			return E_NOTIMPL;
	}
	STDMETHOD(SetData)(FORMATETC* pformatetc ,
		STGMEDIUM* pmedium ,
		BOOL fRelease ) {
			m_fromat = *pformatetc;
			m_stgmed = *pmedium;

			return S_OK;
	}
	STDMETHOD(EnumFormatEtc)(DWORD dwDirection ,
		IEnumFORMATETC** ppenumFormatEtc ) {
			return E_NOTIMPL;
	}
	STDMETHOD(DAdvise)(FORMATETC *pformatetc,
		DWORD advf,
		IAdviseSink *pAdvSink,
		DWORD *pdwConnection) {
			return E_NOTIMPL;
	}
	STDMETHOD(DUnadvise)(DWORD dwConnection) {
		return E_NOTIMPL;
	}
	STDMETHOD(EnumDAdvise)(IEnumSTATDATA **ppenumAdvise) {
		return E_NOTIMPL;
	}
	// Some Other helper functions
	//
	void SetBitmap(HBITMAP hBitmap);
	IOleObject *GetOleObject(IOleClientSite *pOleClientSite,
		IStorage *pStorage);
};

void CImageDataObject::InsertBitmap(IRichEditOle* pRichEditOle, HBITMAP hBitmap)
{
	SCODE sc;

	// Get the image data object
	//
	CImageDataObject *pods = new CImageDataObject;
	LPDATAOBJECT lpDataObject;
	pods->QueryInterface(IID_IDataObject, (void **)&lpDataObject);

	pods->SetBitmap(hBitmap);

	// Get the RichEdit container site
	//
	IOleClientSite *pOleClientSite;	
	pRichEditOle->GetClientSite(&pOleClientSite);

	// Initialize a Storage Object
	//
	IStorage *pStorage;	

	LPLOCKBYTES lpLockBytes = NULL;
	sc = ::CreateILockBytesOnHGlobal(NULL, TRUE, &lpLockBytes);
	if (sc != S_OK)
		exit(EXIT_FAILURE);
	//AfxThrowOleException(sc);
	assert(lpLockBytes != NULL);
	
	sc = ::StgCreateDocfileOnILockBytes(lpLockBytes,
		STGM_SHARE_EXCLUSIVE|STGM_CREATE|STGM_READWRITE, 0, &pStorage);
	if (sc != S_OK)
	{
		
		//verify(lpLockBytes->Release() == 0);
		lpLockBytes = NULL;
		exit(EXIT_FAILURE);
		//AfxThrowOleException(sc);
	}
	assert(pStorage != NULL);

	// The final ole object which will be inserted in the richedit control
	//
	IOleObject *pOleObject; 
	pOleObject = pods->GetOleObject(pOleClientSite, pStorage);

	// all items are "contained" -- this makes our reference to this object
	//  weak -- which is needed for links to embedding silent update.
	OleSetContainedObject(pOleObject, TRUE);

	// Now Add the object to the RichEdit 
	//
	REOBJECT reobject;
	ZeroMemory(&reobject, sizeof(REOBJECT));
	reobject.cbStruct = sizeof(REOBJECT);
	
	CLSID clsid;
	sc = pOleObject->GetUserClassID(&clsid);
	if (sc != S_OK)
		exit(EXIT_FAILURE);
		//AfxThrowOleException(sc);

	reobject.clsid = clsid;
	reobject.cp = REO_CP_SELECTION;
	reobject.dvaspect = DVASPECT_CONTENT;
	reobject.poleobj = pOleObject;
	reobject.polesite = pOleClientSite;
	reobject.pstg = pStorage;

	// Insert the bitmap at the current location in the richedit control
	//
	pRichEditOle->InsertObject(&reobject);

	// Release all unnecessary interfaces
	//
	pOleObject->Release();
	pOleClientSite->Release();
	pStorage->Release();
	lpDataObject->Release();
}


IOleObject *CImageDataObject::GetOleObject(IOleClientSite *pOleClientSite, IStorage *pStorage)
{
	assert(m_stgmed.hBitmap);

	SCODE sc;
	IOleObject *pOleObject;
	sc = ::OleCreateStaticFromData(this, IID_IOleObject, OLERENDER_FORMAT, 
			&m_fromat, pOleClientSite, pStorage, (void **)&pOleObject);
	if (sc != S_OK)
		exit(EXIT_FAILURE);
		//AfxThrowOleException(sc);
	return pOleObject;
}

void CImageDataObject::SetBitmap(HBITMAP hBitmap)
{
	assert(hBitmap);

	STGMEDIUM stgm;
	stgm.tymed = TYMED_GDI;					// Storage medium = HBITMAP handle		
	stgm.hBitmap = hBitmap;
	stgm.pUnkForRelease = NULL;				// Use ReleaseStgMedium

	FORMATETC fm;
	fm.cfFormat = CF_BITMAP;				// Clipboard format = CF_BITMAP
	fm.ptd = NULL;							// Target Device = Screen
	fm.dwAspect = DVASPECT_CONTENT;			// Level of detail = Full content
	fm.lindex = -1;							// Index = Not applicaple
	fm.tymed = TYMED_GDI;					// Storage medium = HBITMAP handle

	this->SetData(&fm, &stgm, TRUE);		
}

// Step 4: the Window Procedure
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	//HDC         hDC;
	//PAINTSTRUCT Ps1;
	COLORREF    clrBlack = RGB(0, 0, 0);
	RECT        Recto = { 10, 220, 210, 400 };
	COLORREF    clrWhite = RGB(255, 255, 255);
	//COLORREF	clrGrey = RGB(0xff, 0xff, 0x00); 

	HDC hDC, MemDCExercising;
	PAINTSTRUCT Ps;
	HBITMAP bmpExercising, bmp2;
	LPRICHEDITOLE m_pRichEditOle;


	switch(msg)
	{
	case WM_CREATE :
		{
			LoadLibrary(TEXT("Riched32.dll"));
			char pMessage[] = "With data on over 60,000 companies covering 300 years,GFD offers a unique source";
			
			HWND hwndEdit = CreateWindowEx(
				0, "RICHEDIT",   // predefined class 
				NULL,         // no window title 
				WS_CHILD | WS_VISIBLE | WS_VSCROLL | 
				ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL  | WS_BORDER , 
				250,
				50, 
				500, 
				350,   // set size in WM_SIZE message 
				hwnd,         // parent window 
				(HMENU) ID_EDITCHILD,   // edit control ID 
				(HINSTANCE) GetWindowLong(hwnd, GWL_HINSTANCE), 
				NULL);



			if(hwndEdit == NULL)
				goto error_statement;
			else{
				HFONT hFont=CreateFont(18,0,0,0,FW_DONTCARE,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_OUTLINE_PRECIS,
					CLIP_DEFAULT_PRECIS,ANTIALIASED_QUALITY, VARIABLE_PITCH,TEXT("Trebuchet MS"));

				SendMessage(hwndEdit,             // Handle of edit control
					WM_SETFONT,         // Message to change the font
					(WPARAM) hFont,     // handle of the font
					MAKELPARAM(TRUE, 0) // Redraw text
					);

				SendMessage(hwndEdit, WM_SETTEXT , 0, (LPARAM)pMessage); 
				//BOOL value = Edit_Enable(hwndEdit,TRUE);
				//SendMessage(hwndEdit, EM_SETREADONLY ,1 ,1);

				SendMessage(hwndEdit,  EM_GETOLEINTERFACE, 0, (LPARAM)&m_pRichEditOle);
				bmp2 = LoadBitmap(GetModuleHandle(NULL), MAKEINTRESOURCE(ID_MY_BITMAP));
				CImageDataObject::InsertBitmap(m_pRichEditOle, bmp2);
				
				
			}

			char pMessage3[] = "Update Package";
			HWND hwndEdit3 = CreateWindowEx(
				0, "EDIT",   // predefined class 
				NULL,         // no window title 
				WS_CHILD | WS_VISIBLE  | 
				ES_LEFT | ES_MULTILINE , 
				250,
				30, 
				500, 
				20,   // set size in WM_SIZE message 
				hwnd,         // parent window 
				(HMENU) ID_EDITCHILD,   // edit control ID 
				(HINSTANCE) GetWindowLong(hwnd, GWL_HINSTANCE), 
				NULL);



			if(hwndEdit3 == NULL)
				goto error_statement;
			else{
				HFONT hFont=CreateFont(18,0,0,0,FW_DONTCARE | FW_BOLD,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_OUTLINE_PRECIS,
					CLIP_DEFAULT_PRECIS,ANTIALIASED_QUALITY, VARIABLE_PITCH,TEXT("Trebuchet MS"));

				SendMessage(hwndEdit3,             // Handle of edit control
					WM_SETFONT,         // Message to change the font
					(WPARAM) hFont,     // handle of the font
					MAKELPARAM(TRUE, 0) // Redraw text
					);
				SendMessage(hwndEdit3, WM_SETTEXT , 0, (LPARAM)pMessage3); 
				SendMessage(hwndEdit3, EM_SETREADONLY ,1 ,1);
			}

			char pMessage2[] = "Important: At the end of this installation a system reboot will be required. Do not install other applications while executing Dell Update Packages.";
			HWND hwndEdit2 = CreateWindowEx(
				0, "EDIT",   // predefined class 
				NULL,         // no window title 
				WS_CHILD | WS_VISIBLE  | 
				ES_LEFT | ES_MULTILINE , 
				250,
				400, 
				500, 
				60,   // set size in WM_SIZE message 
				hwnd,         // parent window 
				(HMENU) ID_EDITCHILD,   // edit control ID 
				(HINSTANCE) GetWindowLong(hwnd, GWL_HINSTANCE), 
				NULL);



			if(hwndEdit2 == NULL)
				goto error_statement;
			else{
				HFONT hFont=CreateFont(18,0,0,0,FW_DONTCARE | FW_BOLD,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_OUTLINE_PRECIS,
					CLIP_DEFAULT_PRECIS,ANTIALIASED_QUALITY, VARIABLE_PITCH,TEXT("Trebuchet MS"));

				SendMessage(hwndEdit2,             // Handle of edit control
					WM_SETFONT,         // Message to change the font
					(WPARAM) hFont,     // handle of the font
					MAKELPARAM(TRUE, 0) // Redraw text
					);
				SendMessage(hwndEdit2, WM_SETTEXT , 0, (LPARAM)pMessage2); 
				SendMessage(hwndEdit2, EM_SETREADONLY ,1 ,1);
			}

			// This is for Install button
			if(!CreateWindow(TEXT("button"), 
				TEXT("Install"),
				WS_VISIBLE | WS_CHILD ,
				565, 
				500, 
				75, 
				23, 
				hwnd,
				//NULL,
				reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_BUTTON_INSTALL)), 
				NULL, 
				NULL))
				goto error_statement;

			// This is for Cancel button
			if(!CreateWindow(TEXT("button"), 
				TEXT("Cancel"),
				WS_VISIBLE | WS_CHILD ,
				675, 
				500, 
				75, 
				23, 
				hwnd,  
				//NULL,
				reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_BUTTON_CANCEL)), 
				NULL, 
				NULL))
				goto error_statement;

			/*HBITMAP hImage = (HBITMAP)LoadImage(NULL, 
			"C:\\Users\\Ananya_Jana\\Documents\\New_Version_FrameWire\\Assets\\darwin.bmp",
			IMAGE_BITMAP,
			50,
			50,
			LR_LOADFROMFILE);*/


			char pMessage1[] = "Hello, World!";
			HWND hwndEdit1 = CreateWindowEx(
				0, "EDIT",   // predefined class 
				NULL,         // no window title 
				WS_CHILD | WS_VISIBLE  | 
				ES_CENTER | ES_MULTILINE  , 
				10,
				220, 
				100, 
				100,   // set size in WM_SIZE message 
				hwnd,         // parent window 
				(HMENU) ID_EDITCHILD,   // edit control ID 
				(HINSTANCE) GetWindowLong(hwnd, GWL_HINSTANCE), 
				NULL);

			if(hwndEdit1 == NULL)
				goto error_statement;

			else{
				HFONT hFont=CreateFont(18,0,0,0,FW_DONTCARE,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_OUTLINE_PRECIS,
					CLIP_DEFAULT_PRECIS,ANTIALIASED_QUALITY, VARIABLE_PITCH,TEXT("Trebuchet MS"));

				SendMessage(hwndEdit1,             // Handle of edit control
					WM_SETFONT,         // Message to change the font
					(WPARAM) hFont,     // handle of the font
					MAKELPARAM(TRUE, 0) // Redraw text
					);
				SendMessage(hwndEdit1, WM_SETTEXT , 0, (LPARAM)pMessage1); 
				SendMessage(hwndEdit1, EM_SETREADONLY ,1 ,1);
				OutputDebugString((LPCSTR)"hello");
			}

		}

		break;

	case WM_COMMAND:	
		switch(HIWORD(wParam))
		{
		case BN_CLICKED:
			{
				switch(LOWORD(wParam)){
				case ID_BUTTON_INSTALL:
					{

					}
					break;

				case ID_BUTTON_CANCEL: case ID_BUTTON_CLOSE:
					{
						DestroyWindow(hwnd);
					}
					break;

				}
				break;
			}

		}
		break;
	case WM_PAINT:
		hDC = BeginPaint(hwnd, &Ps);

		// Load the bitmap from the resource
		bmpExercising = LoadBitmap(GetModuleHandle(NULL), MAKEINTRESOURCE(ID_DELL_BITMAP));
		// Create a memory device compatible with the above DC variable
		MemDCExercising = CreateCompatibleDC(hDC);
		// Select the new bitmap
		SelectObject(MemDCExercising, bmpExercising);

		// Copy the bits from the memory DC into the current dc
		//BitBlt(hDC, 10, 10, 450, 400, MemDCExercising, 0, 0, SRCCOPY);
		BitBlt(hDC, 40, 90, 100, 100, MemDCExercising, 0, 0, SRCCOPY);

		// Restore the old bitmap
		DeleteDC(MemDCExercising);
		DeleteObject(bmpExercising);
		MoveToEx(hDC, 5, 480, NULL);
		LineTo(hDC, 778, 480);
		EndPaint(hwnd, &Ps);
		
		break;
		
		/*
		hDC = BeginPaint(hwnd, &Ps);
		SetTextColor(hDC, clrBlack);
		//SetBkColor(hDC, clrWhite);
		//SetBkColor(hDC, TRANSPARENT);
		ExtTextOut(hDC, 10, 10, NULL,
			&Recto, "(C) 2003-2012", 13, NULL);
		EndPaint(hwnd, &Ps);
		break;
		*/
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
error_statement: {
	DWORD errCode = GetLastError();
	char* err;
	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
		NULL,
		errCode,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // default language
		(LPTSTR) &err,
		0,
		NULL);
	MessageBox(hwnd,err,"Error!",0);
				 }
}


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
	LPSTR lpCmdLine, int nCmdShow)
{
	WNDCLASSEX wc;
	HWND hwnd;
	MSG Msg;

	//Step 1: Registering the Window Class
	wc.cbSize        = sizeof(WNDCLASSEX);
	wc.style         = 0;
	wc.lpfnWndProc   = WndProc;
	wc.cbClsExtra    = 0;
	wc.cbWndExtra    = 0;
	wc.hInstance     = hInstance;			
	wc.hIcon         = LoadIcon(NULL, IDI_APPLICATION);
	wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW);
	//wc.hbrBackground = NULL;
	wc.lpszMenuName  = NULL;
	wc.lpszClassName = g_szClassName;
	wc.hIconSm       = LoadIcon(NULL, IDI_WINLOGO);

	// Registering our window class
	if(!RegisterClassEx(&wc))
	{
		MessageBox(NULL, "Window Registration Failed!", "Error!",
			MB_ICONEXCLAMATION | MB_OK);
		return 0;
	}

	// Step 2: Creating the Window
	hwnd = CreateWindowEx(
		WS_EX_CLIENTEDGE,
		g_szClassName,
		"WireFrame",
		//WS_OVERLAPPEDWINDOW,
		WS_VISIBLE | WS_POPUP | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_OVERLAPPEDWINDOW,
		X_COORD, Y_COORD, 800, 600,
		NULL, NULL, hInstance, NULL);

	if(hwnd == NULL)
	{
		MessageBox(NULL, "Window Creation Failed!", "Error!",
			MB_ICONEXCLAMATION | MB_OK);
		return 0;
	}
	//HWND re=CreateWindowA("RICHEDIT","text",WS_BORDER|WS_CHILD|WS_VISIBLE|ES_MULTILINE,10,10,300,300,hwnd,0,hInstance,0);
	// The window becomes visible after this
	ShowWindow(hwnd, nCmdShow);
	UpdateWindow(hwnd); 

	// Step 3: Loop to retrieve all messages to the fn
	while(GetMessage(&Msg, NULL, 0, 0) > 0)
	{
		TranslateMessage(&Msg);
		DispatchMessage(&Msg);
	}
	return Msg.wParam;
}

