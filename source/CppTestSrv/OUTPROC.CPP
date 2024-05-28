/*
	This code is based on example code to the book:
		Inside COM
		by Dale E. Rogerson
		Microsoft Press 1997
		ISBN 1-57231-349-8
*/

#include <windows.h>
#include <strsafe.h>
#include <cstdlib>

#include "CUnknown.h"
#include "CFactory.h"

///////////////////////////////////////////////////////////
//
// Function to retrieving the Last-Error Code.
// Copied from:
// https://learn.microsoft.com/en-us/windows/win32/Debug/retrieving-the-last-error-code
//
void ErrorExit(LPTSTR lpszFunction) 
{ 
    // Retrieve the system error message for the last-error code

    LPVOID lpMsgBuf;
    LPVOID lpDisplayBuf;
    DWORD dw = GetLastError(); 

    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | 
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL,
        dw,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
        (LPTSTR) &lpMsgBuf,
        0, NULL );

    // Display the error message and exit the process

    lpDisplayBuf = (LPVOID)LocalAlloc(LMEM_ZEROINIT, 
        (lstrlen((LPCTSTR)lpMsgBuf) + lstrlen((LPCTSTR)lpszFunction) + 40) * sizeof(TCHAR)); 
    StringCchPrintf((LPTSTR)lpDisplayBuf, 
        LocalSize(lpDisplayBuf) / sizeof(TCHAR),
        TEXT("%s failed with error %d: %s"), 
        lpszFunction, dw, lpMsgBuf); 
    MessageBox(NULL, (LPCTSTR)lpDisplayBuf, TEXT("Error"), MB_OK); 

    LocalFree(lpMsgBuf);
    LocalFree(lpDisplayBuf);
}

///////////////////////////////////////////////////////////
//
// Outproc.cpp
//   - the component server
//
HWND g_hWndListBox = NULL ;

BOOL InitWindow(int nCmdShow) ;
extern "C" LONG APIENTRY MainWndProc(HWND hWnd,
                                     UINT message,
                                     WPARAM wParam,
                                     LPARAM lParam) ;


//
// WinMain procedure
//
extern "C" int WINAPI WinMain(HINSTANCE hInstance, 
                              HINSTANCE hPrevInstance,
                              LPSTR lpCmdLine, 
                              int nCmdShow)
{
	// Controls whether UI is shown or not
	BOOL bUI = TRUE ;

	// If TRUE, don't loop.
	BOOL bExit = FALSE ;

	// Initialize the COM Library.
	HRESULT hr = OleInitialize(NULL) ;
	if (FAILED(hr))
	{
		return 0 ;
	}
   
   
	// Get Thread ID.
	CFactory::s_dwThreadID = ::GetCurrentThreadId() ;
	CFactory::s_hModule = hInstance ;

	// Read the command line.
	char szTokens[] = "-/" ;

	char* szToken = strtok(lpCmdLine, szTokens) ; 
	while (szToken != NULL)
	{
		if (_stricmp(szToken, "UnregServer") == 0)
		{
			CFactory::UnregisterAll() ;
			// MessageBox(0,L"Successfully removed COM-server registration.", L"server.exe /UnregServer", MB_OK) ;
			// We are done, so exit.
			bExit = TRUE ;
			bUI = FALSE ;
		}
		else if (_stricmp(szToken, "RegServer") == 0)
		{
			CFactory::RegisterAll() ;
			// MessageBox(0,L"Successfully registered COM-server", L"server.exe /RegServer", MB_OK) ;
			// We are done, so exit.
			bExit = TRUE ;
			bUI = FALSE ;
		}
		else if (_stricmp(szToken, "Embedding") == 0)
		{
			// Don't display a window if we are embedded.
			bUI = FALSE ;
			break ;
		}
		szToken = strtok(NULL, szTokens) ;
	}

	// If the user started us, then show UI.
	if (bUI)
	{
		if (!InitWindow(nCmdShow))
		{
			// Exit since we can't show UI.
			bExit = TRUE ;
		}
		else
		{
			::InterlockedIncrement(&CFactory::s_cServerLocks) ;
		}
	}

	if (!bExit)
	{
		// Register all of the class factories.
		CFactory::StartFactories() ;

		// Wait for shutdown.
		MSG msg ;
		while (::GetMessage(&msg, 0, 0, 0))
		{
			::DispatchMessage(&msg) ;
		}

		// Unregister the class factories.
		CFactory::StopFactories() ;
	}

	// Uninitialize the COM Library.
	OleUninitialize() ;
	return 0 ;
}


//
// Initialize window
//
BOOL InitWindow(int nCmdShow) 
{
	// Fill in window class structure with parameters
	// that describe the main window.
	WNDCLASS wcListview ;
	wcListview.style = 0 ;                     
	wcListview.lpfnWndProc = (WNDPROC)MainWndProc ; 
	wcListview.cbClsExtra = 0 ;              
	wcListview.cbWndExtra = 0 ;              
	wcListview.hInstance = CFactory::s_hModule ;
	wcListview.hIcon = ::LoadIcon(NULL, IDI_APPLICATION) ;
	wcListview.hCursor = ::LoadCursor(NULL, IDC_ARROW) ;
	wcListview.hbrBackground = static_cast<HBRUSH>(::GetStockObject(WHITE_BRUSH)) ; 
	wcListview.lpszMenuName = NULL ;  
	wcListview.lpszClassName = L"MyServerWinClass" ;

	BOOL bResult = ::RegisterClass(&wcListview) ;
	if (!bResult)
	{
		return bResult ;
	}

	HWND hWndMain ;

	hWndMain = ::CreateWindow(
		L"MyServerWinClass",
		L"Component Server", 
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT,
		CW_USEDEFAULT, CW_USEDEFAULT,
		NULL,               
		NULL,               
		CFactory::s_hModule,          
		NULL) ;

	// If window could not be created, return "failure".
	if (!hWndMain)
	{
		ErrorExit(TEXT("CreateWindow"));
		return FALSE ;
	}

	// Make the window visible; update its client area;
	// and return "success".
	::ShowWindow(hWndMain, nCmdShow) ;
	::UpdateWindow(hWndMain) ;
	return TRUE ;
}

//
// Main window procedure
//
extern "C" LONG APIENTRY MainWndProc(
	HWND hWnd,                // window handle
	UINT message,             // type of message
	WPARAM wParam,              // additional information
	LPARAM lParam)              // additional information
{
	DWORD dwStyle ;

	switch (message) 
	{
	case WM_CREATE:
		{
			// Get size of main window
			CREATESTRUCT* pcs = (CREATESTRUCT*) lParam ;

			// Create a listbox for output.
			g_hWndListBox = ::CreateWindow(
				L"LISTBOX",
				NULL, 
				WS_CHILD | WS_VISIBLE | LBS_USETABSTOPS
					| WS_VSCROLL | LBS_NOINTEGRALHEIGHT,
				0, 0, pcs->cx, pcs->cy,
				hWnd,               
				NULL,               
				CFactory::s_hModule,          
				NULL) ;
			if (g_hWndListBox  == NULL)
			{
				// Listbox not created.
				::MessageBox(NULL,
				             L"Listbox not created!",
				             NULL,
				             MB_OK) ;
				return -1 ;
			}
		}
		break ;

	case WM_SIZE:
		::MoveWindow(g_hWndListBox, 0, 0,
			LOWORD(lParam), HIWORD(lParam), TRUE) ;
		break;

	case WM_DESTROY:          // message: window being destroyed
		if (CFactory::CanUnloadNow() == S_OK)
		{
			// Only post the quit message, if there is
			// no one using the program.
			::PostQuitMessage(0) ;
		}
		break ;

	case WM_CLOSE:
		// Decrement the lock count.
		::InterlockedDecrement(&CFactory::s_cServerLocks) ;

		// The list box is going away.
		g_hWndListBox = NULL ;

		//Fall through 
	default:
		return (DefWindowProc(hWnd, message, wParam, lParam)) ;
	}
	return 0 ;
}
