#include "stdafx.h"
#include <guiddef.h>
#include "JavaScriptFunction.h"

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

// {CE0A5C2E-BEB2-4579-91C6-307D6145B189}
MIDL_DEFINE_GUID(IID, IID_IJavaScriptFunction, 0XCE0A5C2E, 0XBEB2, 0X4579, 0X91, 0XC6, 0X30, 0X7D, 0X61, 0X45, 0XB1, 0X89);
