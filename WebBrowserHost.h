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

// These constants are not defined cause we set _WIN32_IE to
// _WIN32_IE_IE60SP2, so that IE7 and later features are not
// available in our code, so that backward compatibility is fine.
// But still, we only use this constants features for security
// manager, and that does not make any functionality broken.

#ifndef URLACTION_LOOSE_XAML
#define URLACTION_LOOSE_XAML						0x00002402
#endif // !URLACTION_LOOSE_XAML
#ifndef URLACTION_MANAGED_SIGNED
#define URLACTION_MANAGED_SIGNED					0x00002001
#endif // !URLACTION_MANAGED_SIGNED
#ifndef URLACTION_MANAGED_UNSIGNED
#define URLACTION_MANAGED_UNSIGNED					0x00002004
#endif // !URLACTION_MANAGED_UNSIGNED
#ifndef URLACTION_WINDOWS_BROWSER_APPLICATIONS
#define URLACTION_WINDOWS_BROWSER_APPLICATIONS		0x00002400
#endif // !URLACTION_WINDOWS_BROWSER_APPLICATIONS
#ifndef URLACTION_WINFX_SETUP
#define URLACTION_WINFX_SETUP						0x00002600
#endif // !URLACTION_WINFX_SETUP
#ifndef URLACTION_XPS_DOCUMENTS
#define URLACTION_XPS_DOCUMENTS						0x00002401
#endif // !URLACTION_XPS_DOCUMENTS
#ifndef URLACTION_ALLOW_AUDIO_VIDEO
#define URLACTION_ALLOW_AUDIO_VIDEO					0x00002701
#endif // !URLACTION_ALLOW_AUDIO_VIDEO
#ifndef URLACTION_ALLOW_STRUCTURED_STORAGE_SNIFFING
#define URLACTION_ALLOW_STRUCTURED_STORAGE_SNIFFING	0x00002703
#endif // !URLACTION_ALLOW_STRUCTURED_STORAGE_SNIFFING
#ifndef URLACTION_ALLOW_XDOMAIN_SUBFRAME_RESIZE
#define URLACTION_ALLOW_XDOMAIN_SUBFRAME_RESIZE		0x00001408
#endif // !URLACTION_ALLOW_XDOMAIN_SUBFRAME_RESIZE
#ifndef URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE
#define URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE	0x00002107
#endif // !URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE
#ifndef URLACTION_SHELL_PREVIEW
#define URLACTION_SHELL_PREVIEW						0x0000180F
#endif // !URLACTION_SHELL_PREVIEW
#ifndef URLACTION_SHELL_REMOTEQUERY
#define URLACTION_SHELL_REMOTEQUERY					0x0000180E
#endif // !URLACTION_SHELL_REMOTEQUERY
#ifndef URLACTION_SHELL_SECURE_DRAGSOURCE
#define URLACTION_SHELL_SECURE_DRAGSOURCE			0x0000180D
#endif // !URLACTION_SHELL_SECURE_DRAGSOURCE
#ifndef URLACTION_ALLOW_APEVALUATION
#define URLACTION_ALLOW_APEVALUATION				0x00002301
#endif // !URLACTION_ALLOW_APEVALUATION
#ifndef URLACTION_LOWRIGHTS
#define URLACTION_LOWRIGHTS							0x00002500
#endif // !URLACTION_LOWRIGHTS
#ifndef URLACTION_ALLOW_ACTIVEX_FILTERING
#define URLACTION_ALLOW_ACTIVEX_FILTERING			0x00002702
#endif // !URLACTION_ALLOW_ACTIVEX_FILTERING
#ifndef URLACTION_DOTNET_USERCONTROLS
#define URLACTION_DOTNET_USERCONTROLS				0x00002005
#endif // !URLACTION_DOTNET_USERCONTROLS
#ifndef URLACTION_FEATURE_DATA_BINDING
#define URLACTION_FEATURE_DATA_BINDING              0x00002106
#endif // !URLACTION_FEATURE_DATA_BINDING
#ifndef URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE
#define URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE  0x00002107
#endif // !URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE
#ifndef URLACTION_SCRIPT_XSSFILTER
#define URLACTION_SCRIPT_XSSFILTER                  0x00001409
#endif // !URLACTION_SCRIPT_XSSFILTER


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

	STDMETHOD(ShowContextMenu)(DWORD /*dwID*/, POINT FAR* /*ppt*/, IUnknown* /*pcmdTarget*/, IDispatch* /*pdispObject*/) {
		return S_OK;// CAxHostWindow::ShowContextMenu(dwID, ppt, pcmdTarget, pdispObject);
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
	STDMETHOD(SetSecuritySite)(
		/* [unique][in] */ __RPC__in_opt IInternetSecurityMgrSite * /*pSite*/) {
		return INET_E_DEFAULT_ACTION;
	}

	STDMETHOD(GetSecuritySite)(
		/* [out] */ __RPC__deref_out_opt IInternetSecurityMgrSite ** /*ppSite*/) {
		return INET_E_DEFAULT_ACTION;
	};

	STDMETHOD(MapUrlToZone)(
		/* [in] */ __RPC__in LPCWSTR pwszUrl,
		/* [out] */ __RPC__out DWORD * pdwZone,
		/* [in] */ DWORD /*dwFlags*/) {
		_CRT_UNUSED(pwszUrl);
		*pdwZone = 0;
		return S_OK;
	};

	STDMETHOD(GetSecurityId)(
		/* [in] */ __RPC__in LPCWSTR pwszUrl,
		/* [size_is][out] */ __RPC__out_ecount_full(*pcbSecurityId) BYTE * pbSecurityId,
		/* [out][in] */ __RPC__inout DWORD * pcbSecurityId,
		/* [in] */ DWORD_PTR /*dwReserved*/) {
		_CRT_UNUSED(pwszUrl);
#define MY_SECURITY_DOMAIN "file:"
		int cbSecurityDomain = strlen(MY_SECURITY_DOMAIN);
		if (*pcbSecurityId >= MAX_SIZE_SECURITY_ID) {

			memset(pbSecurityId, 0, *pcbSecurityId);
#pragma warning(disable:4996)
			strcpy((char*)pbSecurityId, MY_SECURITY_DOMAIN);
#pragma warning(default:4996)

			// Last 4 bytes are <URLZONE> and then 3 zeros.
			pbSecurityId[cbSecurityDomain + 1] = URLZONE_LOCAL_MACHINE; // ==0
			pbSecurityId[cbSecurityDomain + 2] = 0;
			pbSecurityId[cbSecurityDomain + 3] = 0;
			pbSecurityId[cbSecurityDomain + 4] = 0;

			*pcbSecurityId = (DWORD)cbSecurityDomain + 4; // plus the 4 bytes from above.
		}

		return S_OK;
//		return INET_E_DEFAULT_ACTION;
	};

	STDMETHOD(ProcessUrlAction)(
		/* [in] */ __RPC__in LPCWSTR /*pwszUrl*/,
		/* [in] */ DWORD dwAction,
		/* [size_is][out] */ __RPC__out_ecount_full(cbPolicy) BYTE * pPolicy,
		/* [in] */ DWORD /*cbPolicy*/,
		/* [unique][in] */ __RPC__in_opt BYTE * /*pContext*/,
		/* [in] */ DWORD /*cbContext*/,
		/* [in] */ DWORD /*dwFlags*/,
		/* [in] */ DWORD /*dwReserved*/) {

		// You can't set all actions to ALLOW, as it sometimes works the other way,
		// an ALLOW could set restrictions we don't want to. We have to check and set
		// each flag separately.

		// URLACTION flags: http://msdn.microsoft.com/en-us/library/ms537178(VS.85).aspx
		// Default settings: http://msdn.microsoft.com/en-us/library/ms537186(VS.85).aspx
		// URLPOLICY flags: http://msdn.microsoft.com/en-us/library/ms537179(VS.85).aspx

		// Not implemented because of different reasons:
		// URLACTION_ACTIVEX_OVERRIDE_DOMAINLIST // ie9, not sure what this does
		// URLACTION_INPRIVATE_BLOCKING // no idea

		switch (dwAction) {
		case URLACTION_ACTIVEX_CONFIRM_NOOBJECTSAFETY:
		case URLACTION_ACTIVEX_OVERRIDE_DATA_SAFETY:
		case URLACTION_ACTIVEX_OVERRIDE_SCRIPT_SAFETY:
		case URLACTION_FEATURE_BLOCK_INPUT_PROMPTS:
		case URLACTION_SCRIPT_OVERRIDE_SAFETY:
		case URLACTION_SHELL_EXTENSIONSECURITY:
		case URLACTION_ACTIVEX_NO_WEBOC_SCRIPT:
		case URLACTION_ACTIVEX_OVERRIDE_OBJECT_SAFETY:
		case URLACTION_ACTIVEX_OVERRIDE_OPTIN:
		case URLACTION_ACTIVEX_OVERRIDE_REPURPOSEDETECTION:
		case URLACTION_ACTIVEX_RUN:
		case URLACTION_ACTIVEX_SCRIPTLET_RUN:
		case URLACTION_ACTIVEX_DYNSRC_VIDEO_AND_ANIMATION:
		case URLACTION_ALLOW_RESTRICTEDPROTOCOLS:
		case URLACTION_AUTOMATIC_ACTIVEX_UI:
		case URLACTION_AUTOMATIC_DOWNLOAD_UI:
		case URLACTION_BEHAVIOR_RUN:
		case URLACTION_CLIENT_CERT_PROMPT:
		case URLACTION_COOKIES:
		case URLACTION_COOKIES_ENABLED:
		case URLACTION_COOKIES_SESSION:
		case URLACTION_COOKIES_SESSION_THIRD_PARTY:
		case URLACTION_COOKIES_THIRD_PARTY:
		case URLACTION_CROSS_DOMAIN_DATA:
		case URLACTION_DOWNLOAD_SIGNED_ACTIVEX:
		case URLACTION_DOWNLOAD_UNSIGNED_ACTIVEX:
		case URLACTION_FEATURE_DATA_BINDING:
		case URLACTION_FEATURE_FORCE_ADDR_AND_STATUS:
		case URLACTION_FEATURE_MIME_SNIFFING:
		case URLACTION_FEATURE_SCRIPT_STATUS_BAR:
		case URLACTION_FEATURE_WINDOW_RESTRICTIONS:
		case URLACTION_FEATURE_ZONE_ELEVATION:
		case URLACTION_HTML_FONT_DOWNLOAD:
		case URLACTION_HTML_INCLUDE_FILE_PATH:
		case URLACTION_HTML_JAVA_RUN:
		case URLACTION_HTML_META_REFRESH:
		case URLACTION_HTML_MIXED_CONTENT:
		case URLACTION_HTML_SUBFRAME_NAVIGATE:
		case URLACTION_HTML_SUBMIT_FORMS:
		case URLACTION_HTML_SUBMIT_FORMS_FROM:
		case URLACTION_HTML_SUBMIT_FORMS_TO:
		case URLACTION_HTML_USERDATA_SAVE:
		case URLACTION_LOOSE_XAML:
		case URLACTION_MANAGED_SIGNED:
		case URLACTION_MANAGED_UNSIGNED:
		case URLACTION_SCRIPT_JAVA_USE:
		case URLACTION_SCRIPT_PASTE:
		case URLACTION_SCRIPT_RUN:
		case URLACTION_SCRIPT_SAFE_ACTIVEX:
		case URLACTION_SHELL_EXECUTE_HIGHRISK:
		case URLACTION_SHELL_EXECUTE_LOWRISK:
		case URLACTION_SHELL_EXECUTE_MODRISK:
		case URLACTION_SHELL_FILE_DOWNLOAD:
		case URLACTION_SHELL_INSTALL_DTITEMS:
		case URLACTION_SHELL_MOVE_OR_COPY:
		case URLACTION_SHELL_VERB:
		case URLACTION_SHELL_WEBVIEW_VERB:
		case URLACTION_WINDOWS_BROWSER_APPLICATIONS:
		case URLACTION_WINFX_SETUP:
		case URLACTION_XPS_DOCUMENTS:
		case URLACTION_ALLOW_AUDIO_VIDEO: // ie9
		case URLACTION_ALLOW_STRUCTURED_STORAGE_SNIFFING: // ie9
		case URLACTION_ALLOW_XDOMAIN_SUBFRAME_RESIZE: // ie7
		case URLACTION_FEATURE_CROSSDOMAIN_FOCUS_CHANGE: // ie7
		case URLACTION_SHELL_ENHANCED_DRAGDROP_SECURITY:
		case URLACTION_SHELL_PREVIEW: // win7
		case URLACTION_SHELL_REMOTEQUERY: // win7
		case URLACTION_SHELL_RTF_OBJECTS_LOAD: // ie6sp2
		case URLACTION_SHELL_SECURE_DRAGSOURCE: // ie7
		// case URLACTION_SHELL_SHELLEXECUTE: // ie6sp2, value the same as URLACTION_SHELL_EXECUTE_HIGHRISK
		case URLACTION_DOTNET_USERCONTROLS: // ie8, probably registry only
			*pPolicy = URLPOLICY_ALLOW;
			return S_OK;

		case URLACTION_CHANNEL_SOFTDIST_PERMISSIONS:
			//*pPolicy = URLPOLICY_CHANNEL_SOFTDIST_AUTOINSTALL;
			*pPolicy = URLPOLICY_ALLOW;
			return S_OK;

		case URLACTION_JAVA_PERMISSIONS:
			//*pPolicy = URLPOLICY_JAVA_LOW;
			*pPolicy = URLPOLICY_ALLOW;
			return S_OK;

		case URLACTION_CREDENTIALS_USE:
			//*pPolicy = URLPOLICY_CREDENTIALS_SILENT_LOGON_OK;
			*pPolicy = URLPOLICY_ALLOW;
			return S_OK;

		case URLACTION_ALLOW_APEVALUATION: // Phishing filter.
		case URLACTION_LOWRIGHTS: // Vista Protected Mode.
		case URLACTION_SHELL_POPUPMGR:
		case URLACTION_SCRIPT_XSSFILTER:
		case URLACTION_ACTIVEX_TREATASUNTRUSTED:
		case URLACTION_ALLOW_ACTIVEX_FILTERING: // ie9
			*pPolicy = URLPOLICY_DISALLOW;
			return S_FALSE;
		}

		return INET_E_DEFAULT_ACTION;
	};

	STDMETHOD(QueryCustomPolicy)(
		/* [in] */ __RPC__in LPCWSTR /*pwszUrl*/,
		/* [in] */ __RPC__in REFGUID /*guidKey*/,
		/* [size_is][size_is][out] */ __RPC__deref_out_ecount_full_opt(*pcbPolicy) BYTE ** /*ppPolicy*/,
		/* [out] */ __RPC__out DWORD * /*pcbPolicy*/,
		/* [in] */ __RPC__in BYTE * /*pContext*/,
		/* [in] */ DWORD /*cbContext*/,
		/* [in] */ DWORD /*dwReserved*/) {
		return INET_E_DEFAULT_ACTION;
	};

	STDMETHOD(SetZoneMapping)(
		/* [in] */ DWORD /*dwZone*/,
		/* [in] */ __RPC__in LPCWSTR /*lpszPattern*/,
		/* [in] */ DWORD /*dwFlags*/) {
		return E_NOTIMPL;
	};

	STDMETHOD(GetZoneMappings)(
		/* [in] */ DWORD /*dwZone*/,
		/* [out] */ __RPC__deref_out_opt IEnumString ** /*ppenumString*/,
		/* [in] */ DWORD /*dwFlags*/) {
		return E_NOTIMPL;
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

