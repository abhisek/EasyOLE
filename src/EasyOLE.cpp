#include <windows.h>
#include <ole2.h>

#include "EasyOLE.h"

//#define HAVE_DEBUG_UTILS
#ifdef HAVE_DEBUG_UTILS
#include "Debug.h"
#endif

static HRESULT EasyOleFindIID(
	IDispatch		*pIDispatch,
	LPOLESTR		*pName,
	IID				*pIID,
	ITypeInfo		**ppTypeInfo)
{
	ITypeInfo		*pTypeInfo;
	ITypeLib		*pTypeLib;

	return E_FAIL;
}

HRESULT EasyOleRegisterEventHandler(
	IDispatch		*pIDispatch,
	LPOLESTR		*pOleEvent,
	EASY_OLE_EVH	*Handler)
{
	return E_FAIL;
}

// Thanks: http://www.codeproject.com/Articles/34998/MS-Office-OLE-Automation-Using-C
static HRESULT EasyOleInternalDispatch(
		INT			nType,
		VARIANT		*pvResult,
		IDispatch	*pIDispatch,
		LPOLESTR	pOleName,
		INT			cArgs...			
	)
{
	va_list		vMarker;
	HRESULT		hResult;
	DISPPARAMS	dp = {NULL, NULL, 0, 0};
	DISPID		dispNamedId = DISPID_PROPERTYPUT;
	DISPID		dispId;
	VARIANT		*pArgs = NULL;
	EXCEPINFO	execpInfo;
	UINT		puArgErr = 0;

	if(!pIDispatch)
		return E_FAIL;

	hResult = pIDispatch->GetIDsOfNames(IID_NULL, &pOleName, 1, 
		LOCALE_USER_DEFAULT, &dispId);
	if(FAILED(hResult)) {
#ifdef HAVE_DEBUG_UTILS
		DebugWriteMsg("OleInvoke: Failed to get DispId");
#endif
		return hResult;
	}

	va_start(vMarker, cArgs);

	pArgs = new VARIANT[cArgs + 1];

	// Populate pArgs in reverse order as expected by OLE
	for(INT i = (cArgs - 1); i >= 0; i--) {
		pArgs[i] = va_arg(vMarker, VARIANT);
	}

	//DebugWriteMsg("Arg Count: %d Arg0.lVal: 0x%08x", cArgs, pArgs[0].lVal);

	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	if(nType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispNamedId;
	}

	ZeroMemory(&execpInfo, sizeof(execpInfo));

	if(pvResult)
		VariantInit(pvResult);

	hResult = pIDispatch->Invoke(dispId, IID_NULL, 
		LOCALE_SYSTEM_DEFAULT, nType, &dp, pvResult, &execpInfo, &puArgErr);

	if(FAILED(hResult)) {
#ifdef HAVE_DEBUG_UTILS
		DebugWriteMsg("OleInvoke: Exception [scode: 0x%lX wcode: %u puArgErr: %u] GLE: 0x%08x", 
			execpInfo.scode, execpInfo.wCode, puArgErr, GetLastError());
#endif

		SysFreeString(execpInfo.bstrDescription);
		SysFreeString(execpInfo.bstrHelpFile);
		SysFreeString(execpInfo.bstrSource);
	}

	va_end(vMarker);
	delete [] pArgs;

	return hResult;
}

HRESULT EasyOleMethodCall(
	IDispatch	*pIDispatch,
	VARIANT		*pvResult,
	LPOLESTR	pOleName,
	INT			cArgs,
	VARIANT		lpArg1,
	VARIANT		lpArg2,
	VARIANT		lpArg3,
	VARIANT		lpArg4,
	VARIANT		lpArg5
	)
{
	return EasyOleInternalDispatch(DISPATCH_METHOD, 
		pvResult, pIDispatch, pOleName, cArgs,
		lpArg1, lpArg2, lpArg3, lpArg4, lpArg5);
}

HRESULT EasyOlePropertyGet(
	IDispatch	*pIDispatch,
	VARIANT		*pvResult,
	LPOLESTR	pOleName,
	INT			cArgs,
	VARIANT		lpArg1,
	VARIANT		lpArg2,
	VARIANT		lpArg3,
	VARIANT		lpArg4,
	VARIANT		lpArg5
	)
{
	return EasyOleInternalDispatch(DISPATCH_PROPERTYGET,
		pvResult, pIDispatch, pOleName, cArgs,
		lpArg1, lpArg2, lpArg3, lpArg4, lpArg5);
}

HRESULT EasyOlePropertyPut(
	IDispatch	*pIDispatch,
	VARIANT		*pvResult,
	LPOLESTR	pOleName,
	INT			cArgs,
	VARIANT		lpArg1,
	VARIANT		lpArg2,
	VARIANT		lpArg3,
	VARIANT		lpArg4,
	VARIANT		lpArg5
	)
{
	return EasyOleInternalDispatch(DISPATCH_PROPERTYPUT,
		pvResult, pIDispatch, pOleName, cArgs,
		lpArg1, lpArg2, lpArg3, lpArg4, lpArg5);
}

VARIANT EasyOleStringToVariant(LPOLESTR pString)
{
	VARIANT v;

	VariantInit(&v);
	V_VT(&v) = VT_BSTR;
	V_BSTR(&v) = pString;

	return v;
}

VARIANT EasyOleIntToVariant(INT i)
{
	VARIANT v;

	VariantInit(&v);
	V_VT(&v) = VT_I4;
	V_I4(&v) = i;

	return v;
}

VOID EasyOleReleaseObject(IDispatch *pIDispatch, LPVOID lpObject)
{
	
}

VOID EasyOleReleaseInstance(IDispatch *pIDispatch)
{
	pIDispatch->Release();
}

HRESULT EasyOleCreateInstance(LPCOLESTR lpszProgID, IDispatch **ppIDispatch)
{
	HRESULT		hResult;
	CLSID		clsid;
	LPVOID		p;

	hResult = CLSIDFromProgID(lpszProgID, &clsid);
	if(FAILED(hResult))
		return hResult;

	hResult = CoCreateInstance((REFCLSID) clsid, 0, 
		CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER, 
		(REFCLSID) IID_IDispatch, &p);

	if(FAILED(hResult))
		return hResult;

	*ppIDispatch = (IDispatch*) p;

	return hResult;
}

BOOL EasyOleMessageLoopOnce()
{
	MSG		msg;

	if(PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);

		return TRUE;
	}

	return FALSE;
}

VOID EasyOleMessageLoop()
{
	MSG msg;

	while(EasyOleMessageLoopOnce())
		;
}

BOOL EasyOleInit(DWORD dwInitType)
{
	HRESULT		hResult;

	if(!dwInitType)
		dwInitType = COINIT_MULTITHREADED;

	hResult = CoInitializeEx(0, dwInitType);
	if(FAILED(hResult)) {
		return FALSE;
	}

	return TRUE;
}