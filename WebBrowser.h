//
// WebBrowser.h: creates IWebBrowser2 control
//
// Copyright (C) 2012 Hong Jen Yee (PCMan) <pcman.tw@gmail.com>
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

#pragma once

#include <atlbase.h>
#include <atlwin.h>
#include <ExDispid.h>

#include "WebBrowserHost.h"

const int _nDispatchID = 1;

class CWebBrowserDelegate
{
public:
	virtual void OnBeforeNavigate(const wchar_t* /*uri*/, bool* /*pCancel*/) {}
	virtual void OnNewWindow(const wchar_t* /*uri*/, const wchar_t* /*ref*/, bool* pCancel, IDispatch** ppDisp) { *pCancel = false; *ppDisp = nullptr; }
	virtual void OnNavigateComplete(const wchar_t* /*uri*/) {}
	virtual void OnDocumentComplete(const wchar_t* /*uri*/) {}
	virtual void OnTitleChange(const wchar_t* /*title*/) {}
	virtual void OnSetStatusText(const wchar_t* /*text*/) {}
	virtual void OnCommandStatusChange(bool /*cangoback*/, bool /*cangofwd*/) {}
	virtual void OnBrowserWindowClosing(bool* /*pCancel*/) {}
	virtual bool OnBrowserWindowMessage(UINT /*message*/, WPARAM /*wparam*/, LPARAM /*lparam*/) { return false; }
	virtual void OnRedirectXDomainBlocked(const wchar_t* /*startUrl*/, const wchar_t* /*redirectUrl*/, const wchar_t* /*frame*/) {}

};

class /* ATL_NO_VTABLE */ CWebBrowser:
	public CWindowImpl<CWebBrowser, CAxWindow>,
	public IDispEventImpl<_nDispatchID, CWebBrowser>,
	public CMessageFilter
{
public:
	CWebBrowser();
	~CWebBrowser();

	DECLARE_WND_SUPERCLASS(_T("WebBrowserHost"), GetWndClassName())

	HWND Create(HWND parentWnd, RECT& rc, DWORD dwStyle) {
		HWND ret = CWindowImpl<CWebBrowser, CAxWindow>::Create(parentWnd, &rc, _T(""), dwStyle);
		return ret;
	}

	bool GetCanBack() {
		return m_CanBack;
	}

	bool GetCanForward() {
		return m_CanForward;
	}

	CComPtr<IWebBrowser2> GetIWebBrowser2() {
		return m_pIWebBrowser2;
	}

	static CWebBrowser* FromHwnd(HWND hwnd);

	void SetDelegate(CWebBrowserDelegate * pDelegate) {
		m_pDelegate = pDelegate;
	}

	HWND GetActiveXWnd()
	{
		ATLASSERT(::IsWindow(m_hWnd));

		HWND hActiveXWnd = NULL;
		HWND hIEControlWnd = NULL;
		HWND hShellhWnd = ::FindWindowEx(m_hWnd, NULL, _T("Shell Embedding"), NULL);
		if (hShellhWnd)
		{
			hShellhWnd = ::FindWindowEx(hShellhWnd, NULL, _T("Shell DocObject View"), NULL);
			if (hShellhWnd)
			{
				hIEControlWnd = ::FindWindowEx(hShellhWnd, NULL, _T("Internet Explorer_Server"), NULL);
				if (hIEControlWnd)
				{
					hActiveXWnd = GetBottomChildWnd(hIEControlWnd);
					if (hActiveXWnd == hIEControlWnd)
						hActiveXWnd = NULL;
				}
			}
		}
		return hActiveXWnd;
	}


private:
	BEGIN_MSG_MAP(CWebBrowser)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
	END_MSG_MAP()

	LRESULT OnCreate(UINT uMsg, WPARAM wParam , LPARAM lParam, BOOL& bHandled);
	LRESULT OnDestroy(UINT uMsg, WPARAM wParam , LPARAM lParam, BOOL& bHandled);
	// LRESULT OnParentNotify(UINT uMsg, WPARAM wParam , LPARAM lParam, BOOL& bHandled);

	// override this to do cleanup
	virtual void OnFinalMessage(HWND hwnd);

	// WindowProc for innermost Internet_Explorer_Server window.
	static LRESULT InnerWndProc(HWND hwnd, UINT message, WPARAM wparam, LPARAM lparam);
	LRESULT HandleInnerWndProc(UINT message, WPARAM wparam, LPARAM lparam);
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_SINK_MAP(CWebBrowser)
		SINK_ENTRY(_nDispatchID, DISPID_BEFORENAVIGATE2, OnBeforeNavigate2)
		SINK_ENTRY(_nDispatchID, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete2)
		SINK_ENTRY(_nDispatchID, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete)	
		SINK_ENTRY(_nDispatchID, DISPID_NEWWINDOW3, OnNewWindow3)
		SINK_ENTRY(_nDispatchID, DISPID_STATUSTEXTCHANGE, OnStatusTextChange)
		SINK_ENTRY(_nDispatchID, DISPID_TITLECHANGE, OnTitleChange)
		SINK_ENTRY(_nDispatchID, DISPID_WINDOWCLOSING, OnWindowClosing)
		SINK_ENTRY(_nDispatchID, DISPID_COMMANDSTATECHANGE, OnCommandStateChange)
		SINK_ENTRY(_nDispatchID, DISPID_REDIRECTXDOMAINBLOCKED, OnRedirectXDomainBlocked)
	END_SINK_MAP()

	void __stdcall OnNewWindow3(IDispatch **ppDisp, VARIANT_BOOL *Cancel, long flags, BSTR bstrUrlContext, BSTR bstrUrl);
	void __stdcall OnBeforeNavigate2(IDispatch *pDisp, VARIANT *url, VARIANT *Flags, VARIANT *TargetFrameName, VARIANT *PostData, VARIANT *Headers, VARIANT_BOOL *Cancel);
	void __stdcall OnNavigateComplete2(IDispatch* pDisp,  VARIANT* URL);
	void __stdcall OnDocumentComplete(IDispatch* pDisp, VARIANT* URL);
	void __stdcall OnStatusTextChange(BSTR Text);
	void __stdcall OnTitleChange(BSTR Text);
	void __stdcall OnWindowClosing(VARIANT_BOOL IsChildWindow, VARIANT_BOOL *Cancel);
	void __stdcall OnCommandStateChange(long Command, VARIANT_BOOL Enable);
	void __stdcall OnRedirectXDomainBlocked(IDispatch *pDisp, VARIANT *StartUrl, VARIANT *RedirectUrl, VARIANT *Frame, VARIANT *StatusCode);
	
	HWND GetBottomChildWnd(HWND hParentWnd)
	{
		HWND hRetWindow = hParentWnd;
		HWND hWndChild = ::GetWindow(hRetWindow, GW_CHILD);
		while (hWndChild != NULL)
		{
			hRetWindow = hWndChild;
			hWndChild = ::GetWindow(hRetWindow, GW_CHILD);
		}

		return hRetWindow;
	}

	bool Navigate(LPCTSTR lpstrURL, int nFlag = 0, LPCTSTR pTargetName = 0, LPCTSTR pPostData = 0, LPCTSTR pHeader = 0)
	{
		bool bSuccess = false;

		if (m_pIWebBrowser2 != NULL)
		{
			CComVariant vtURL(lpstrURL);
			CComVariant vtFlag(nFlag);
			CComVariant vtTargetName;
			if (pTargetName != NULL)
			{
				vtTargetName = pTargetName;
			}
			CComVariant vtPostData;
			if (pPostData != NULL)
			{
				vtPostData = pPostData;
			}
			CComVariant vtHeader;
			if (pHeader != NULL)
			{
				vtHeader = pHeader;
			}

			HRESULT hRet = m_pIWebBrowser2->Navigate2(&vtURL, &vtFlag, &vtTargetName, &vtPostData, &vtHeader);
			bSuccess = SUCCEEDED(hRet);

		}
		return bSuccess;
	}

	bool Refresh()
	{
		bool bSuccess = false;
		if (m_pIWebBrowser2 != NULL)
		{
			HRESULT hRet = m_pIWebBrowser2->Refresh();
			bSuccess = SUCCEEDED(hRet);
		}
		return bSuccess;
	}

	bool GetJScript(CComPtr<IDispatch>& spDisp)
	{
		CComPtr<IDispatch> spdispDoc;
		HRESULT hRet = m_pIWebBrowser2->get_Document(&spdispDoc);
		if (FAILED(hRet))
			return SUCCEEDED(hRet);
		CComQIPtr<IHTMLDocument2> htmlDoc = spdispDoc;

		hRet = htmlDoc->get_Script(&spDisp);
		ATLASSERT(SUCCEEDED(hRet));

		return SUCCEEDED(hRet);
	}

	void ExportToJScript(const ATL::CString &strFunc, UINT argCount, ExportFunction::FunctionDelegate function)
	{
		CComPtr<IAxWinHostWindow> pAxWindow = (IAxWinHostWindow*)GetWindowLongPtr(GWLP_USERDATA);
		if (pAxWindow)
		{
			IJavaScriptFunction * javaScriptFunction;
			pAxWindow->QueryInterface(IID_IJavaScriptFunction, (void**)&javaScriptFunction);
			if (javaScriptFunction)
			{
				javaScriptFunction->ExportFunctionToJS(strFunc, argCount, function);
				javaScriptFunction->Release();
			}
		}
	}

	bool CallJScript(const ATL::CString &strFunc, CComVariant* pVarResult = NULL)
	{
		ATL::CSimpleArray<ATL::CString> params;
		return CallJScript(strFunc, params, pVarResult);
	}
	bool CallJScript(const ATL::CString &strFunc, const ATL::CString & param1, CComVariant* pVarResult = NULL)
	{
		ATL::CSimpleArray<ATL::CString> params;
		params.Add(param1);
		return CallJScript(strFunc, params, pVarResult);
	}
	bool CallJScript(const ATL::CString &strFunc, 
		const ATL::CString & param1, const ATL::CString & param2,
		CComVariant* pVarResult = NULL)
	{
		ATL::CSimpleArray<ATL::CString> params;
		params.Add(param1);
		params.Add(param2);
		return CallJScript(strFunc, params, pVarResult);
	}
	bool CallJScript(const ATL::CString &strFunc, 
		const ATL::CString & param1, const ATL::CString & param2, const ATL::CString & param3,
		CComVariant* pVarResult = NULL)
	{
		ATL::CSimpleArray<ATL::CString> params;
		params.Add(param1);
		params.Add(param2);
		params.Add(param3);
		return CallJScript(strFunc, params, pVarResult);
	}

	bool CallJScript(const ATL::CString &strFunc, const ATL::CSimpleArray<ATL::CString>& paramArray, CComVariant* pVarResult = NULL)
	{
		//Getting IDispatch for Java Script objects
		CComPtr<IDispatch> spScript;
		if (!GetJScript(spScript))
		{
			return false;
		}
		//Find dispid for given function in the object
		CComBSTR bstrMember(strFunc);
		DISPID dispid = NULL;
		HRESULT hr = spScript->GetIDsOfNames(IID_NULL, &bstrMember, 1,
			LOCALE_SYSTEM_DEFAULT, &dispid);
		if (FAILED(hr))
		{
			return false;
		}

		const int arraySize = paramArray.GetSize();
		//Putting parameters  
		DISPPARAMS dispparams;
		memset(&dispparams, 0, sizeof dispparams);
		dispparams.cArgs = arraySize;
		dispparams.rgvarg = new VARIANT[dispparams.cArgs];
		dispparams.cNamedArgs = 0;

		for (int i = 0; i < arraySize; i++)
		{
			CComBSTR bstr = paramArray[arraySize - 1 - i]; // back reading
			bstr.CopyTo(&dispparams.rgvarg[i].bstrVal);
			dispparams.rgvarg[i].vt = VT_BSTR;
		}
		EXCEPINFO excepInfo;
		memset(&excepInfo, 0, sizeof excepInfo);
		CComVariant vaResult;
		UINT nArgErr = (UINT)-1;  // initialize to invalid arg
								  //Call JavaScript function         
		hr = spScript->Invoke(dispid, IID_NULL, 0,
			DISPATCH_METHOD, &dispparams,
			&vaResult, &excepInfo, &nArgErr);
		delete[] dispparams.rgvarg;
		if (FAILED(hr))
		{
			return false;
		}
		if (pVarResult)
		{
			*pVarResult = vaResult;
		}

		return true;
	}

private:
	CComPtr<IWebBrowser2> m_pIWebBrowser2;

	HWND m_hInnerWnd;
	WNDPROC m_OldInnerWndProc;
	CComQIPtr<IOleInPlaceActiveObject> m_pInPlaceActiveObject;
	bool m_CanBack;
	bool m_CanForward;
	bool m_Destroyed;

	static int browserCount;
	static ATOM winPropAtom;

	CWebBrowserDelegate * m_pDelegate;
};
