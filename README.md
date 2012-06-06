EasyOLE: C/C++ OLE Automation Client Library
============================================

This library is written to ease the development of OLE Automation Clients in C/C++ handling the internal not-so-friendly aspects of COM interfaces. This library is far from complete and a lot of enhancements are in the pipeline.

EasyOLE in general supports 3 core operation on an OLE Automation Interface:

* Property Set
* Property Get
* Method Call

OLE needs to be initialized in the current process before any calls can be made:

```C
EasyOleInit(0);
```


In order to issue calls to an OLE Automation Interface (IDispatch), client programs need to obtain a pointer to its IDispatch interface:

```C
hResult = EasyOleCreateInstance(TEXT("SKYPE4COM.Skype"), &pSkypeDispatch);
if(FAILED(hResult) {
  [...]
}
```

Once an IDispatch interface is obtained, various operation supported by the interface can be performed. [OLE-COM Object Viewer](http://msdn.microsoft.com/en-us/library/windows/desktop/ms688269\(v=vs.85\).aspx) from Microsoft SDK Tools can be used to enumerate the COM interfaces and view corresponding TypeLib information.

Example Client
---------------

```C
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
```

TODO
-----

* OLE Event Handlers