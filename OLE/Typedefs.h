#ifndef WORD_TYPEDEFS_H
#define WORD_TYPEDEFS_H

#include <ole2.h>
#include "Enums.h"
#include <memory>
#include <string>

#include "StdAfx.h"
#include "conversions.h"


#define SafeRelease(p)      if (p) { p->Release(); p = 0; }

#define Validation3(fnName, param, hr)                                         \
{                                                                              \
    if (FAILED(hr)) {                                                          \
        reportFailure(fnName, param, hr);                                      \
        return;                                                                \
    }                                                                          \
}

#define Validation2(fnName, hr)    Validation3(fnName, std::string(""), hr)

#endif
