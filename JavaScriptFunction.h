#pragma once

#include <atlbase.h>
#include <atlhost.h>
#include <atlcoll.h>
#include <atlstr.h>

const DISPID DISPID_BASE = 1234;

struct ExportFunction
{
	DISPID funcId;
	ATL::CString funcName;
	UINT argCount;
	typedef std::function<void(DISPPARAMS*, VARIANT*)> FunctionDelegate;
	FunctionDelegate func;
};

EXTERN_C const IID IID_IJavaScriptFunction;
MIDL_INTERFACE("CE0A5C2E-BEB2-4579-91C6-307D6145B189")
IJavaScriptFunction : public IDispatch 
{
public:
	// 使用CAtlArray或者CSimpleArray数据结构添加多个元素时function对象内存错乱，原因未知待查
	std::vector<ExportFunction> m_exportFunctions; 
	CAtlMap<ATL::CString, DISPID, ATL::CStringElementTraitsI<ATL::CString>> m_functionIdMaps;

	//virtual ULONG STDMETHODCALLTYPE AddRef(void)
	//{
	//	return 1;
	//}

	//virtual ULONG STDMETHODCALLTYPE Release(void)
	//{
	//	return 1;
	//}

	//STDMETHOD(QueryInterface)(
	//	/* [in] */ REFIID riid,
	//	/* [iid_is][out] */ __RPC__deref_out void __RPC_FAR *__RPC_FAR *ppvObject)
	//{
	//	if (ppvObject == NULL)
	//		return E_POINTER;
	//	*ppvObject = NULL;
	//	if (InlineIsEqualGUID(riid, IID_IUnknown))
	//	{
	//		*ppvObject = static_cast<IUnknown *>(this);
	//		return S_OK;
	//	}
	//	if (InlineIsEqualGUID(riid, IID_IDispatch))
	//	{
	//		*ppvObject = static_cast<IDispatch *>(this);
	//		return S_OK;
	//	}
	//	if (InlineIsEqualGUID(riid, IID_IJavaScriptFunction))
	//	{
	//		*ppvObject = static_cast<IJavaScriptFunction *>(this);
	//		return S_OK;
	//	}
	//	return E_NOINTERFACE;
	//}

	STDMETHOD(GetTypeInfoCount)(
		/* [out] */ __RPC__out UINT * pctinfo) {
		*pctinfo = 0;
		return S_OK;
	}

	STDMETHOD(GetTypeInfo)(
		/* [in] */ UINT /*iTInfo*/,
		/* [in] */ LCID /*lcid*/,
		/* [out] */ __RPC__deref_out_opt ITypeInfo ** /*ppTInfo*/) {
		return E_FAIL;
	}

	STDMETHOD(GetIDsOfNames)(
		_In_ REFIID /*riid*/,
		_In_reads_(cNames) _Deref_pre_z_ LPOLESTR* rgszNames,
		_In_range_(0, 16384) UINT cNames,
		LCID /*lcid*/,
		_Out_ DISPID* rgdispid)
	{
		if (cNames == 0 || rgszNames == NULL || rgszNames[0] == NULL || rgdispid == NULL)
			return E_INVALIDARG;
		auto func = m_functionIdMaps.Lookup(rgszNames[0]);
		if (func != NULL)
		{
			*rgdispid = func->m_value;
			return S_OK;
		}

		return E_NOTIMPL;
	}
	STDMETHOD(Invoke)(
		_In_ DISPID dispidMember,
		_In_ REFIID /*riid*/,
		_In_ LCID /*lcid*/,
		_In_ WORD /*wFlags*/,
		_In_ DISPPARAMS* pdispparams,
		_Out_opt_ VARIANT* pvarResult,
		_Out_opt_ EXCEPINFO* /*pexcepinfo*/,
		_Out_opt_ UINT* /*puArgErr*/)
	{
		if (dispidMember >= DISPID_BASE && dispidMember < DISPID_BASE + (DISPID)m_exportFunctions.size())
		{
			int funcIndex = dispidMember - DISPID_BASE;
			const ExportFunction & func = m_exportFunctions[funcIndex];
			_ASSERT(pdispparams->cArgs == func.argCount);
			for (UINT i = 0; i < pdispparams->cArgs; ++i)
			{
				if ((pdispparams->rgvarg[i].vt & VT_BSTR) == 0)
				{
					return E_INVALIDARG;
				}
			}
			func.func(pdispparams, pvarResult);
			return S_OK;
		}

		return E_NOTIMPL;
	}

	STDMETHOD(ExportFunctionToJS)(const ATL::CString & functionName, UINT argCount, ExportFunction::FunctionDelegate function) {
		if (functionName == _T("") || function == NULL)
			return E_INVALIDARG;

		ExportFunction func;
		func.funcName = functionName;
		func.argCount = argCount;
		func.func = function;
		func.funcId = DISPID_BASE + m_exportFunctions.size();
		m_exportFunctions.push_back(func);
		m_functionIdMaps[func.funcName] = func.funcId;
		return S_OK;
	}

};