#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <richedit.h>

TCHAR szClassName[] = TEXT("window");

BOOL PrintRTF(HWND hwnd, HDC hdc)
{
	DOCINFO di = { sizeof(di) };
	if (!StartDoc(hdc, &di))
	{
		return FALSE;
	}
	int cxPhysOffset = GetDeviceCaps(hdc, PHYSICALOFFSETX);
	int cyPhysOffset = GetDeviceCaps(hdc, PHYSICALOFFSETY);
	int cxPhys = GetDeviceCaps(hdc, PHYSICALWIDTH);
	int cyPhys = GetDeviceCaps(hdc, PHYSICALHEIGHT);
	// Create "print preview". 
	SendMessage(hwnd, EM_SETTARGETDEVICE, (WPARAM)hdc, cxPhys / 2);
	FORMATRANGE fr;
	fr.hdc = hdc;
	fr.hdcTarget = hdc;
	// Set page rect to physical page size in twips.
	fr.rcPage.top = 0;
	fr.rcPage.left = 0;
	fr.rcPage.right = MulDiv(cxPhys, 1440, GetDeviceCaps(hdc, LOGPIXELSX));
	fr.rcPage.bottom = MulDiv(cyPhys, 1440, GetDeviceCaps(hdc, LOGPIXELSY));
	// Set the rendering rectangle to the pintable area of the page.
	fr.rc.left = cxPhysOffset;
	fr.rc.right = cxPhysOffset + cxPhys;
	fr.rc.top = cyPhysOffset;
	fr.rc.bottom = cyPhysOffset + cyPhys;
	SendMessage(hwnd, EM_SETSEL, 0, (LPARAM)-1);          // Select the entire contents.
	SendMessage(hwnd, EM_EXGETSEL, 0, (LPARAM)&fr.chrg);  // Get the selection into a CHARRANGE.
	BOOL fSuccess = TRUE;
	// Use GDI to print successive pages.
	while (fr.chrg.cpMin < fr.chrg.cpMax && fSuccess)
	{
		fSuccess = StartPage(hdc) > 0;
		if (!fSuccess) break;
		int cpMin = (int)SendMessage(hwnd, EM_FORMATRANGE, TRUE, (LPARAM)&fr);
		if (cpMin <= fr.chrg.cpMin)
		{
			fSuccess = FALSE;
			break;
		}
		fr.chrg.cpMin = cpMin;
		fSuccess = EndPage(hdc) > 0;
	}
	SendMessage(hwnd, EM_FORMATRANGE, FALSE, 0);
	if (fSuccess)
	{
		EndDoc(hdc);
	}
	else
	{
		AbortDoc(hdc);
	}
	return fSuccess;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HINSTANCE hRtLib;
	static HWND hEdit;
	static HWND hButton;
	switch (msg) {
	case WM_CREATE:
		hRtLib = LoadLibrary(TEXT("Msftedit.DLL"));
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("印刷..."), WS_VISIBLE | WS_CHILD, 0, 0, 256, 32, hWnd, (HMENU)1000, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit = CreateWindow(TEXT("RICHEDIT50W"), TEXT("http://support.microsoft.com/"), WS_CHILD | WS_VISIBLE | ES_MULTILINE | WS_HSCROLL | WS_VSCROLL | ES_AUTOVSCROLL | ES_NOHIDESEL, 120, 10, 480, 28, hWnd, (HMENU)1001, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
		{
			CHARFORMAT2 cf;
			ZeroMemory(&cf, sizeof(CHARFORMAT2));
			cf.cbSize = sizeof(CHARFORMAT2);
			cf.dwMask = CFM_COLOR | CFM_BACKCOLOR;
			cf.crTextColor = RGB(255,0,0);
			cf.dwEffects = 0;
			lstrcpy(cf.szFaceName, TEXT("メイリオ"));
			cf.yHeight = 200;
			cf.crBackColor = RGB(0,0,255);
			SendMessage(hEdit, EM_SETCHARFORMAT, 0, (LPARAM)&cf);
		}
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == 1000)
		{
			PRINTDLG pd = { 0 };
			DOCINFO di = { 0 };
			pd.lStructSize = sizeof(pd);
			pd.hwndOwner = hWnd;
			pd.hDevMode = NULL;
			pd.hDevNames = NULL;
			pd.Flags = PD_RETURNDC;
			pd.nFromPage = 1;
			pd.nToPage = 1;
			pd.nMinPage = 1;
			pd.nMaxPage = 1;
			pd.nCopies = 1;
			di.cbSize = sizeof(di);
			di.lpszDocName = TEXT("Test Print");
			di.lpszOutput = NULL;
			di.lpszDatatype = NULL;
			di.fwType = 0;
			if (!PrintDlg(&pd))return 0;
			HDC hdc = pd.hDC;
			PrintRTF(hEdit, hdc);
			DeleteDC(hdc);
		}
		break;
	case WM_SIZE:
		MoveWindow(hEdit, 0, 32, LOWORD(lParam), HIWORD(lParam)-32, TRUE);
		break;
	case WM_DESTROY:
		DestroyWindow(hEdit);
		FreeLibrary(hRtLib);
		PostQuitMessage(0);
		break;
	default:
		return(DefWindowProc(hWnd, msg, wParam, lParam));
	}
	return (0L);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		(HBRUSH)(COLOR_WINDOW + 1),
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("RichEdit の内容を色付きで印刷する"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}
