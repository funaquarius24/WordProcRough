#include "conversions.h"

#include <QString>
#include <iostream>

string_t toUtf8( const wstring_t& source )
{
    string_t result;
//    Poco::UnicodeConverter::toUTF8(source, result);
    QString qtString = QString::fromWCharArray( source.c_str() );
    result = qtString.toStdString();
    return result;
}

wstring_t toUtf16( const string_t& source )
{
    wstring_t result;
//    Poco::UnicodeConverter::toUTF8(source, result);
    QString qtString = QString::fromStdString(source);
    result = qtString.toStdWString();
    return result;
}

LPOLESTR toLPOLESTR (const string_t& source )
{
    QString te = source.c_str();

    LPOLESTR strFxn = nullptr;


    te.toWCharArray(strFxn);
    return strFxn;
}

void reportFailure( const string_t& fnName, const string_t& params, HRESULT hr )
{
    char buf[256];

    sprintf_s(buf,
        "%s(%s) failed. hr = 0x%08lx", fnName.c_str(), params.c_str(), hr);

    // TODO: later use logging
    //::MessageBoxA(NULL, buf,  "Error", 0x10010);
    throw std::runtime_error(buf);
}
