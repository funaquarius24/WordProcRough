#include "StdAfx.h"
#include "Font.h"
#include "OLEMethod.h"
#include "Range.h"
#include "conversions.h"


Font::Font( IDispatch* font )
    : font_(font)
{

}

Font::~Font()
{
    SafeRelease(font_);
}

string_t Font::getName() const
{
    return toUtf8(getPropStr(font_, toLPOLESTR("Name")));
}

void Font::setName( const string_t& faceName )
{
    setPropStr(font_, toLPOLESTR("Name"), toUtf16(faceName));
}

void Font::reset()
{
    OLEMethod(DISPATCH_METHOD, NULL, font_, toLPOLESTR("Reset"), 0);
}

void Font::setSize( int sz )
{
    setPropertyInt(font_, toLPOLESTR("Size"), sz);
}

void Font::setColor( COLORREF clr )
{
    setPropertyInt(font_, toLPOLESTR("Color"), clr);
}

void Font::setBold( int bold )
{
    setPropertyInt(font_, toLPOLESTR("Bold"), bold);
}

void Font::setUnderlineColor( COLORREF clr )
{
    setPropertyInt(font_, toLPOLESTR("UnderlineColor"), clr);
}

void Font::setUnderline( int underline )
{
    setPropertyInt(font_, toLPOLESTR("Underline"), underline);
}

void Font::setItalic( int italic )
{
    setPropertyInt(font_, toLPOLESTR("Italic"), italic);
}

int Font::getSize() const
{
    return (int) getPropertyFloat(font_, toLPOLESTR("Size"));
}

int Font::getColor() const
{
    return getPropertyInt(font_, toLPOLESTR("Color"));
}

int Font::getBold() const
{
    return getPropertyInt(font_, toLPOLESTR("Bold"));
}

int Font::getUnderlineColor() const
{
    return getPropertyInt(font_, toLPOLESTR("UnderlineColor"));
}

int Font::getUnderline() const
{
    return getPropertyInt(font_, toLPOLESTR("Underline"));
}

int Font::getItalic() const
{
    return getPropertyInt(font_, toLPOLESTR("Italic"));
}

IDispatch* Font::getIDispatch()
{
    return font_;
}

tFontSp Font::duplicate()
{
    return tFontSp(new Font(getPropertyDispatch(font_, toLPOLESTR("Duplicate"))) );
}

bool Font::haveCommonAttributes(const tRangeSp& r, int& fail)
{
    int i = 0;
    switch (fail) {
    case 0:  if (getSize()      == wdUndefined)               { fail = 0; return false; }
    case 1:  if (getBold()      == wdUndefined)               { fail = 1; return false; }
    case 2:  if (getItalic()    == wdUndefined)               { fail = 2; return false; }
    case 3:  if (getUnderline() == wdUndefined)               { fail = 3; return false; }
    case 4:  if (getColor()     == wdUndefined)               { fail = 4; return false; }
    case 5:  if (r->getHighlightColorIndex() == wdUndefined)  { fail = 5; return false; }
    case 6:  if (getName().empty())                           { fail = 6; return false; }
    }
    fail = 0;
    return true;
}

