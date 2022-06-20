#include "StdAfx.h"
#include "Find.h"
#include "OLEMethod.h"
#include "Range.h"

Find::Find( IDispatch* find )
    : BaseObject(find)
{

}

bool Find::execute()
{
    VARIANT result;
    VariantInit(&result);

    OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("Execute"), 0);

    return (result.bVal == TRUE ? true : false);
}

void Find::setText( const wstring_t& text )
{
    setPropStr(disp_, toLPOLESTR("Text"), text);
}

wstring_t Find::getText() const
{
    return getPropStr(disp_, toLPOLESTR("Text"));
}

void Find::clearFormatting()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("ClearFormatting"), 0);
}

tFontSp Find::getFont() const
{
    return tFontSp(new Font(getPropertyDispatch(disp_, toLPOLESTR("Font"))) );
}
