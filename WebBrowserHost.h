//
// WebBrowserHost.h: Subclass CAxHostWindow to create a specialized
// control host to provide IOleCommandTarget and IDocHostUIHandler.
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
#include <atlhost.h>
#include "JavaScriptFunction.h"

class ATL_NO_VTABLE CWebBrowserHost :
	//public CComCoClass<CWebBrowserHost , &CLSID_NULL>,
	public CAxHostWindow,
	public IInternetSecurityManager,
	public IJavaScriptFunction,
	public IOleCommandTarget
{
protected:
	IInternetSecurityManager* m_pSecurityMgr;

public:
	CWebBrowserHost(void);
	~CWebBrowserHost(void);

	HRESULT FinalConstruct() {
		/* Create Internet Security Manager Object */
		::CoCreateInstance(CLSID_InternetSecurityManager, NULL,
			CLSCTX_INPROC_SERVER, IID_IInternetSecurityManager,
			(void**)&m_pSecurityMgr);

		return S_OK;
	}

	void FinalRelease()
	{
		if (m_pSecurityMgr)
		{
			m_pSecurityMgr->Release();
		}
	}

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	DECLARE_NO_REGISTRY()
	DECLARE_POLY_AGGREGATABLE(CWebBrowserHost)
	DECLARE_GET_CONTROLLING_UNKNOWN()

	BEGIN_SERVICE_MAP(CWebBrowserHost)
		SERVICE_ENTRY(__uuidof(IInternetSecurityManager))
	END_SERVICE_MAP()

	BEGIN_COM_MAP(CWebBrowserHost)
		COM_INTERFACE_ENTRY(IDocHostUIHandler)
		COM_INTERFACE_ENTRY(IOleCommandTarget)
		COM_INTERFACE_ENTRY(IJavaScriptFunction)
		COM_INTERFACE_ENTRY_CHAIN(CAxHostWindow)
	END_COM_MAP()


	// IDocHostUIHandler
	STDMETHOD(GetHostInfo)(DOCHOSTUIINFO FAR* pInfo);
	STDMETHOD(TranslateAccelerator)(LPMSG lpMsg, const GUID FAR* pguidCmdGroup, DWORD nCmdID) {
		return CAxHostWindow::TranslateAccelerator(lpMsg, pguidCmdGroup, nCmdID);
		// return E_NOTIMPL;
	}
	STDMETHOD(GetExternal)(IDispatch** ppDispatch) {
		return CAxHostWindow::GetExternal(ppDispatch);
		// return E_NOTIMPL;
	}

	STDMETHOD(ShowContextMenu)(DWORD dwID, POINT FAR* ppt, IUnknown* pcmdTarget, IDispatch* pdispObject) {
		return CAxHostWindow::ShowContextMenu(dwID, ppt, pcmdTarget, pdispObject);
	}

	STDMETHOD(ShowUI)(DWORD /*dwID*/, IOleInPlaceActiveObject FAR* /*pActiveObject*/,
		IOleCommandTarget FAR* /*pCommandTarget*/,
		IOleInPlaceFrame  FAR* /*pFrame*/,
		IOleInPlaceUIWindow FAR* /*pDoc*/) {
			return S_FALSE;
	}
	STDMETHOD(HideUI)(void) {
		return S_OK;
	}
	STDMETHOD(UpdateUI)(void) {
		return S_OK;
	}
	STDMETHOD(EnableModeless)(BOOL /*fEnable*/) {
		return E_NOTIMPL;
	}
	STDMETHOD(OnDocWindowActivate)(BOOL /*fActivate*/) {
		return E_NOTIMPL;
	}
	STDMETHOD(OnFrameWindowActivate)(BOOL /*fActivate*/) {
		return E_NOTIMPL;
	}
	STDMETHOD(ResizeBorder)(LPCRECT /*prcBorder*/, IOleInPlaceUIWindow FAR* /*pUIWindow*/, BOOL /*fRameWindow*/) {
		return E_NOTIMPL;
	}
	STDMETHOD(GetOptionKeyPath)(LPOLESTR FAR* /*pchKey*/, DWORD /*dw*/) {
		return S_FALSE;
	}
	STDMETHOD(GetDropTarget)(IDropTarget* /*pDropTarget*/, IDropTarget** /*ppDropTarget*/) {
		return E_NOTIMPL;
	}
	STDMETHOD(TranslateUrl)(DWORD /*dwTranslate*/, OLECHAR* /*pchURLIn*/, OLECHAR** ppchURLOut) {
		* ppchURLOut = NULL;
		return S_FALSE;
	}
	STDMETHOD(FilterDataObject)(IDataObject* /*pDO*/, IDataObject** ppDORet) {
		*ppDORet = NULL;
		return S_FALSE;
	}

	// IOleCommandTarget
	STDMETHOD(QueryStatus)(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
	STDMETHOD(Exec)(
		/* [unique][in] */ const GUID *pguidCmdGroup,
		/* [in] */ DWORD nCmdID,
		/* [in] */ DWORD nCmdexecopt,
		/* [unique][in] */ VARIANT *pvaIn,
		/* [unique][out][in] */ VARIANT *pvaOut);


	STDMETHOD(QueryService)(
		REFGUID guidService,
		REFIID riid,
		void __RPC_FAR *__RPC_FAR *ppvObject)
	{
		if (guidService == SID_SInternetSecurityManager &&
			riid == IID_IInternetSecurityManager)
		{
			*ppvObject = dynamic_cast<IInternetSecurityManager*>(this);
			((IInternetSecurityManager*)*ppvObject)->AddRef();
			//ATLTRACE(L"%d\n", hr);
			return S_OK;
		}

		return CAxHostWindow::QueryService(guidService, riid, ppvObject);
	}

	// IInternetSecurityManager
	virtual HRESULT STDMETHODCALLTYPE SetSecuritySite(
		/* [unique][in] */ __RPC__in_opt IInternetSecurityMgrSite * /*pSite*/) {
		return INET_E_DEFAULT_ACTION;
	}

	virtual HRESULT STDMETHODCALLTYPE GetSecuritySite(
		/* [out] */ __RPC__deref_out_opt IInternetSecurityMgrSite ** /*ppSite*/) {
		return INET_E_DEFAULT_ACTION;
	};

	virtual HRESULT STDMETHODCALLTYPE MapUrlToZone(
		/* [in] */ __RPC__in LPCWSTR /*pwszUrl*/,
		/* [out] */ __RPC__out DWORD * /*pdwZone*/,
		/* [in] */ DWORD /*dwFlags*/) {
		return INET_E_DEFAULT_ACTION;
	};

	virtual HRESULT STDMETHODCALLTYPE GetSecurityId(
		/* [in] */ __RPC__in LPCWSTR /*pwszUrl*/,
		/* [size_is][out] */ __RPC__out_ecount_full(*pcbSecurityId) BYTE * /*pbSecurityId*/,
		/* [out][in] */ __RPC__inout DWORD * /*pcbSecurityId*/,
		/* [in] */ DWORD_PTR /*dwReserved*/) {
		return INET_E_DEFAULT_ACTION;
	};

	virtual HRESULT STDMETHODCALLTYPE ProcessUrlAction(
		/* [in] */ __RPC__in LPCWSTR /*pwszUrl*/,
		/* [in] */ DWORD /*dwAction*/,
		/* [size_is][out] */ __RPC__out_ecount_full(cbPolicy) BYTE * pPolicy,
		/* [in] */ DWORD cbPolicy,
		/* [unique][in] */ __RPC__in_opt BYTE * /*pContext*/,
		/* [in] */ DWORD /*cbContext*/,
		/* [in] */ DWORD /*dwFlags*/,
		/* [in] */ DWORD /*dwReserved*/) {
		DWORD dwPolicy = URLPOLICY_ALLOW;
		if (cbPolicy >= sizeof(DWORD))
		{
			*(DWORD*)pPolicy = dwPolicy;
			return S_OK;
		}

		return INET_E_DEFAULT_ACTION;
	};

	virtual HRESULT STDMETHODCALLTYPE QueryCustomPolicy(
		/* [in] */ __RPC__in LPCWSTR /*pwszUrl*/,
		/* [in] */ __RPC__in REFGUID /*guidKey*/,
		/* [size_is][size_is][out] */ __RPC__deref_out_ecount_full_opt(*pcbPolicy) BYTE ** /*ppPolicy*/,
		/* [out] */ __RPC__out DWORD * /*pcbPolicy*/,
		/* [in] */ __RPC__in BYTE * /*pContext*/,
		/* [in] */ DWORD /*cbContext*/,
		/* [in] */ DWORD /*dwReserved*/) {
		return INET_E_DEFAULT_ACTION;
	};

	virtual HRESULT STDMETHODCALLTYPE SetZoneMapping(
		/* [in] */ DWORD /*dwZone*/,
		/* [in] */ __RPC__in LPCWSTR /*lpszPattern*/,
		/* [in] */ DWORD /*dwFlags*/) {
		return INET_E_DEFAULT_ACTION;
	};

	virtual HRESULT STDMETHODCALLTYPE GetZoneMappings(
		/* [in] */ DWORD /*dwZone*/,
		/* [out] */ __RPC__deref_out_opt IEnumString ** /*ppenumString*/,
		/* [in] */ DWORD /*dwFlags*/) {
		return INET_E_DEFAULT_ACTION;
	};

	// IJavaScriptFunction
	STDMETHOD(GetTypeInfoCount)(
		/* [out] */ __RPC__out UINT * pctinfo) {
		return IJavaScriptFunction::GetTypeInfoCount(pctinfo);
	}

	STDMETHOD(GetTypeInfo)(
		/* [in] */ UINT iTInfo,
		/* [in] */ LCID lcid,
		/* [out] */ __RPC__deref_out_opt ITypeInfo ** ppTInfo) {
		return IJavaScriptFunction::GetTypeInfo(iTInfo, lcid, ppTInfo);
	}

	STDMETHOD(GetIDsOfNames)(
		_In_ REFIID riid,
		_In_reads_(cNames) _Deref_pre_z_ LPOLESTR* rgszNames,
		_In_range_(0, 16384) UINT cNames,
		LCID lcid,
		_Out_ DISPID* rgdispid)
	{
		return IJavaScriptFunction::GetIDsOfNames(riid, rgszNames, cNames, lcid, rgdispid);
	}
	STDMETHOD(Invoke)(
		_In_ DISPID dispidMember,
		_In_ REFIID riid,
		_In_ LCID lcid,
		_In_ WORD wFlags,
		_In_ DISPPARAMS* pdispparams,
		_Out_opt_ VARIANT* pvarResult,
		_Out_opt_ EXCEPINFO* pexcepinfo,
		_Out_opt_ UINT* puArgErr)
	{
		return IJavaScriptFunction::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr);
	}

	static HRESULT CWebBrowserHost::AxCreateControlLicEx(LPCOLESTR lpszName, HWND hWnd, IStream* pStream, 
		IUnknown** ppUnkContainer, IUnknown** ppUnkControl, REFIID iidSink, IUnknown* punkSink, BSTR bstrLic);
};

