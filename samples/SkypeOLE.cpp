#include <Windows.h>
#include "../src/EasyOLE.h"

VOID SkypeOleStart()
{
	IDispatch	*pSkypeDispatch;
	IDispatch	*pSkypeClient;
	HRESULT		hResult;
	VARIANT		v1, v2;

	EasyOleInit(0);
	EasyOleCreateInstance(TEXT("SKYPE4COM.Skype"), &pSkypeDispatch);

	// We have to call: SkypeOLE#Client#Start

	// Get Client Property
	hResult = EasyOlePropertyGet(pSkypeDispatch,
		&v1, TEXT("Client"), 0,
		NULL_VARIANT, NULL_VARIANT, NULL_VARIANT, NULL_VARIANT, NULL_VARIANT);

	if(FAILED(hResult))
		return;

	pSkypeClient = V_DISPATCH(&v1);

	// Call Start in Client Dispatch
	V_VT(&v1) = VT_BOOL;
	V_VT(&v2) = VT_BOOL;

	V_BOOL(&v1) = 0;
	V_BOOL(&v2) = 0;

	hResult = EasyOleMethodCall(pSkypeClient, NULL, TEXT("Start"),
		2, v1, v2, NULL_VARIANT, NULL_VARIANT, NULL_VARIANT);

	if(FAILED(hResult)) {
		// Failed to call
	}
}