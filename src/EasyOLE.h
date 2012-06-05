#ifndef _EASY_OLE_H
#define _EASY_OLE_H

#include <Ole2.h>

BOOL EasyOleInit(DWORD dwInitType);
HRESULT EasyOleCreateInstance(LPCOLESTR lpszProgID, IDispatch **ppIDispatch);

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
	);

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
	);

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
	);

VOID EasyOleMessageLoop();
BOOL EasyOleMessageLoopOnce();

static	VARIANT			g_NULL_VARIANT;
#define NULL_VARIANT	(g_NULL_VARIANT)
typedef VOID (EASY_OLE_EVH)(LPVOID lpData);

#endif