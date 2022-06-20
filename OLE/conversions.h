#ifndef CONVERSIONS_H
#define CONVERSIONS_H

#include "Typedefs.h"

string_t   toUtf8( const wstring_t& source );
wstring_t   toUtf16( const string_t& source );
void reportFailure( const string_t& fnName, const string_t& params, HRESULT hr );
LPOLESTR toLPOLESTR (const string_t& source );

#endif // CONVERSIONS_H
