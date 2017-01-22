//
// WebBrowser.cpp: creates IWebBrowser2 control
//
// Copyright (C) 2012 - 2013 Hong Jen Yee (PCMan) <pcman.tw@gmail.com>
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http:www.gnu.org/licenses/>.
//

#include "stdafx.h"
#include "WebBrowser.h"
#include <windowsx.h>

int CWebBrowser::browserCount = 0;
ATOM CWebBrowser::winPropAtom = 0;

CWebBrowser::CWebBrowser():
	m_pIWebBrowser2(NULL),
	m_hInnerWnd(NULL),
	m_CanBack(false),
	m_CanForward(false),
	m_OldInnerWndProc(NULL),
	m_pDelegate(NULL),
	m_Destroyed(false) {
}

CWebBrowser::~CWebBrowser() {
}

// static
CWebBrowser* CWebBrowser::FromHwnd(HWND hwnd) {
	CWebBrowser* pWebBrowser;
	do {
		pWebBrowser = reinterpret_cast<CWebBrowser*>(GetProp(hwnd, reinterpret_cast<LPCTSTR>(winPropAtom)));
		hwnd = ::GetParent(hwnd);
	}while(pWebBrowser == NULL && hwnd != NULL);
	return pWebBrowser;
}

// The message loop of Mozilla does not handle accelertor keys.
// IOleInplaceActivateObject requires MSG be filtered by its TranslateAccellerator() method.
// So we install a hook to do the dirty hack.
// Mozilla message loop is here:
// http://mxr.mozilla.org/mozilla-central/source/widget/src/windows/nsAppShell.cpp
// bool nsAppShell::ProcessNextNativeEvent(bool mayWait)
// It does PeekMessage, TranslateMessage, and then pass the result directly
// to DispatchMessage.
BOOL CWebBrowser::PreTranslateMessage(MSG* pMsg) {
	if ((pMsg->message < WM_KEYFIRST || pMsg->message > WM_KEYLAST) &&
		(pMsg->message < WM_MOUSEFIRST || pMsg->message > WM_MOUSELAST))
		return FALSE;

	BOOL bRet = FALSE;
	// give HTML page a chance to translate this message
	if (pMsg->hwnd == m_hWnd || IsChild(pMsg->hwnd))
		bRet = (BOOL)SendMessage(WM_FORWARDMSG, 0, (LPARAM)pMsg);

	return bRet;
}

LRESULT CWebBrowser::OnCreate(UINT uMsg, WPARAM wParam , LPARAM lParam, BOOL& bHandled) {

	// See ATL source code in atlhost.h
	// static LRESULT CALLBACK AtlAxWindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
	// ATL creates a new CAxHostWindow object when handling WM_CREATE message.
	// We intercept WM_CREATE, and create our CWebBrowserHost object instead.

	CREATESTRUCT* lpCreate = (CREATESTRUCT*)lParam;
	int nCreateSize = 0;
	if (lpCreate && lpCreate->lpCreateParams)
		nCreateSize = *((WORD*)lpCreate->lpCreateParams);
	HGLOBAL h = GlobalAlloc(GHND, nCreateSize);
	CComPtr<IStream> spStream;
	if (h && nCreateSize) {
		BYTE* pBytes = (BYTE*) GlobalLock(h);
		BYTE* pSource = ((BYTE*)(lpCreate->lpCreateParams)) + sizeof(WORD); 
		//Align to DWORD
		//pSource += (((~((DWORD)pSource)) + 1) & 3);
		memcpy(pBytes, pSource, nCreateSize);
		GlobalUnlock(h);
		CreateStreamOnHGlobal(h, TRUE, &spStream);
	}

	IAxWinHostWindow* pAxWindow = NULL;
	CComPtr<IUnknown> spUnk;
	LPCTSTR lpstrName = _T("about:blank");
	HRESULT hRet = CWebBrowserHost::AxCreateControlLicEx(T2COLE(lpstrName), m_hWnd, spStream, &spUnk, NULL, IID_NULL, NULL, NULL);

	if(FAILED(hRet))
		return -1;	// abort window creation
	hRet = spUnk->QueryInterface(__uuidof(IAxWinHostWindow), (void**)&pAxWindow);
	if(FAILED(hRet))
		return -1;	// abort window creation

	SetWindowLongPtr(GWLP_USERDATA, (DWORD_PTR)pAxWindow);
	// continue with DefWindowProc
	LRESULT ret = ::DefWindowProc(m_hWnd, uMsg, wParam, lParam);
	bHandled = TRUE;

	// Now, the Web Browser Active X control is created.
	// Let's do what we want.

	CMessageLoop* pLoop = _Module.GetMessageLoop();
	if(pLoop)
		pLoop->AddMessageFilter(this);

	// Set up event sink
	IUnknown* pUnk = NULL;
	if(SUCCEEDED(QueryControl<IUnknown>(&pUnk))) {

		AtlGetObjectSourceInterface(pUnk, &m_libid, &m_iid, &m_wMajorVerNum, &m_wMinorVerNum);
		DispEventAdvise(pUnk, &DIID_DWebBrowserEvents2);

		pUnk->QueryInterface(IID_IWebBrowser2, (void**)&m_pIWebBrowser2);
		pUnk->Release();

		m_pInPlaceActiveObject = m_pIWebBrowser2; // store the IOleInPlaceActiveObject iface for future use
		m_pIWebBrowser2->put_RegisterAsBrowser(VARIANT_TRUE);
		m_pIWebBrowser2->put_RegisterAsDropTarget(VARIANT_TRUE);
	}

	++browserCount;
	if(browserCount == 1) {
		winPropAtom = GlobalAddAtom(L"WebBrowser::Atom");
	}

	return ret;
}

LRESULT CWebBrowser::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/ , LPARAM /*lParam*/, BOOL& bHandled) {
	if(m_Destroyed == false) {
		// It seems that ATL incorrectly calls OnDestroy twice for one window.
		// So we need to guard OnDestroy with a flag. Otherwise we will uninitilaize things 
		// twice and get cryptic and terrible crashes.

		DispEventUnadvise(m_pIWebBrowser2, &DIID_DWebBrowserEvents2);
		m_pIWebBrowser2.Release();
		m_pInPlaceActiveObject.Release();

		CComPtr<IAxWinHostWindow> pAxWindow = (IAxWinHostWindow*)GetWindowLongPtr(GWLP_USERDATA);
		if (pAxWindow)
			pAxWindow->SetExternalDispatch(NULL);

		if(m_hInnerWnd != NULL) { // if we subclassed the inner most Internet Explorer Server window
			// subclass it back
			::SetWindowLongPtr(m_hInnerWnd, GWL_WNDPROC, (LONG_PTR)m_OldInnerWndProc);
			ATLTRACE("window removed!!\n");
			RemoveProp(m_hInnerWnd, reinterpret_cast<LPCTSTR>(winPropAtom));
			m_hInnerWnd = NULL;
		}

		CMessageLoop* pLoop = _Module.GetMessageLoop();
		if (pLoop)
			pLoop->RemoveMessageFilter(this);


		--browserCount;
		ATLASSERT(browserCount >=0); // this value should never < 0. otherwise it's a bug.

		if(0 == browserCount) {
			GlobalDeleteAtom(winPropAtom);
			winPropAtom = NULL;
		}
		m_Destroyed = true;
	}
	bHandled = TRUE;
	return 0;
}

// NewWindow3 is only available after Window xp SP2
void CWebBrowser::OnNewWindow3(IDispatch ** ppDisp, VARIANT_BOOL * Cancel, long /*flags*/, BSTR bstrUrlContext, BSTR bstrUrl) {
	// ATLTRACE("NewWindow3\n");
	if (m_pDelegate) {
		bool cancel = false;
		m_pDelegate->OnNewWindow(bstrUrl, bstrUrlContext, &cancel, ppDisp);
		*Cancel = cancel ? VARIANT_TRUE : VARIANT_FALSE;
	}
}

void CWebBrowser::OnBeforeNavigate2(IDispatch * /*pDisp*/, VARIANT * URL, VARIANT * /*Flags*/, VARIANT * /*TargetFrameName*/, VARIANT * /*PostData*/, VARIANT * /*Headers*/, VARIANT_BOOL * Cancel) {
	// ATLTRACE("BeforeNavigage\n");
	if (m_pDelegate) {
		bool cancel = false;
		m_pDelegate->OnBeforeNavigate(URL->bstrVal, &cancel);
		*Cancel = cancel ? VARIANT_TRUE : VARIANT_FALSE;
	}
}


void CWebBrowser::OnNavigateComplete2(IDispatch* /*pDisp*/,  VARIANT* URL) {
	// NOTE: dirty hack!! We create the control by loading "about:blank".
	// So soon after the web page "about:blank" is loaded, it's child window
	// which really contains the Internet_Explorer_Server window is created.
	// The first NavigateComplete2 event we got is for "about:blank".
	// So we try to get the inner most child window here
	if(m_hInnerWnd == NULL) {
		m_hInnerWnd = m_hWnd;
		for(;;) {
			HWND child = ::GetWindow(m_hInnerWnd, GW_CHILD);
			if(child == NULL)
				break;
			m_hInnerWnd = child;
		}
		// subclass the window
		m_OldInnerWndProc = (WNDPROC)::SetWindowLongPtr(m_hInnerWnd, GWL_WNDPROC, (LONG_PTR)InnerWndProc);
		SetProp(m_hInnerWnd, reinterpret_cast<LPCTSTR>(winPropAtom), reinterpret_cast<HANDLE>(this));
	}
	// ATLTRACE("NavigateComplete\n");
	if (m_pDelegate) {
		m_pDelegate->OnNavigateComplete(URL->bstrVal);
	}
}

void __stdcall CWebBrowser::OnDocumentComplete(IDispatch* /*pDisp*/, VARIANT* URL)
{
	if (m_pDelegate) {
		m_pDelegate->OnDocumentComplete(URL->bstrVal);
	}
}

void CWebBrowser::OnStatusTextChange(BSTR Text) {
	if (m_pDelegate) {
		m_pDelegate->OnSetStatusText(Text);
	}
}

void CWebBrowser::OnTitleChange(BSTR Text) {
	if (m_pDelegate) {
		m_pDelegate->OnTitleChange(Text);
	}
}

void CWebBrowser::OnWindowClosing(VARIANT_BOOL /*IsChildWindow*/, VARIANT_BOOL *Cancel) {
	// Let's cancel it, so IE won't destroy the web browser control.
	// Then, we call Firefox to close the tab for us.
	*Cancel = VARIANT_TRUE;

	if (m_pDelegate) {
		bool cancel = false;
		m_pDelegate->OnBrowserWindowClosing(&cancel);
		*Cancel = cancel ? VARIANT_TRUE : VARIANT_FALSE;
	}
}

void CWebBrowser::OnCommandStateChange(long Command, VARIANT_BOOL Enable) {
	switch(Command) {
	case CSC_NAVIGATEFORWARD:
		m_CanForward = (Enable == VARIANT_TRUE);
		break;
	case CSC_NAVIGATEBACK:
		m_CanBack = (Enable == VARIANT_TRUE);
		break;
	}
	if (m_pDelegate) {
		m_pDelegate->OnCommandStatusChange(m_CanBack, m_CanForward);
	}
}

void CWebBrowser::OnFinalMessage(HWND /*hwnd*/) {
	/* cleanup when the window is destroyed */
}

// WindowProc for innermost Internet_Explorer_Server window.
LRESULT CWebBrowser::HandleInnerWndProc(UINT message, WPARAM wparam, LPARAM lparam) {
	if (m_pDelegate) {
		if (m_pDelegate->OnBrowserWindowMessage(message, wparam, lparam))
			return 0;
	}
	return CallWindowProc(m_OldInnerWndProc, m_hInnerWnd, message, wparam, lparam);
}

// WindowProc for innermost Internet_Explorer_Server window.
LRESULT CWebBrowser::InnerWndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam) {
	CWebBrowser* pWebBrowser = reinterpret_cast<CWebBrowser*>(GetProp(hwnd, reinterpret_cast<LPCTSTR>(winPropAtom)));
	if(pWebBrowser)
		return pWebBrowser->HandleInnerWndProc(message, wparam, lparam);
	return 0; // This is not possible
}
