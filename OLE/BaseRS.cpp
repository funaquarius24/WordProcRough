#include "StdAfx.h"
#include "BaseRS.h"
#include "OLEMethod.h"
#include "Range.h"


BaseObject::BaseObject( IDispatch* disp )
    : disp_(disp)
{
}

BaseObject::~BaseObject()
{
    SafeRelease(disp_);
}

IDispatch* BaseObject::getIDispatch() const
{
    return disp_;
}


BaseRS::BaseRS( IDispatch* base )
    : BaseObject(base)
{
}

int BaseRS::getStoryLength() const
{
    return getPropertyInt(disp_, toLPOLESTR("StoryLength"));
}

tFontSp BaseRS::getFont() const
{
    return tFontSp(new Font(getPropertyDispatch(disp_, toLPOLESTR("Font"))) );
}

void BaseRS::setFont( const tFontSp& font )
{
    setPropertyDispatch(disp_, font->getIDispatch(), toLPOLESTR("Font"));
}

int BaseRS::setStart( int pos )
{
    return setPropertyInt(disp_, toLPOLESTR("Start"), pos);
}

int BaseRS::getStart() const
{
    return getPropertyInt(disp_, toLPOLESTR("Start"));
}

int BaseRS::setEnd( int pos )
{
    return setPropertyInt(disp_, toLPOLESTR("End"), pos);
}

int BaseRS::getEnd() const
{
    return getPropertyInt(disp_, toLPOLESTR("End"));
}

void BaseRS::setText( const wstring_t& text )
{
    setPropStr(disp_, toLPOLESTR("Text"), text);
}

wstring_t BaseRS::getText() const
{
    return getPropStr(disp_, toLPOLESTR("Text"));
}

void BaseRS::setFormattedText( const tRangeSp& range )
{
    setPropertyDispatch(disp_, range->getIDispatch(), toLPOLESTR("FormattedText"));
}

tRangeSp BaseRS::getFormattedText() const
{
    return tRangeSp(new Range(getPropertyDispatch(disp_, toLPOLESTR("FormattedText"))) );
}

void BaseRS::select()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("Select"), 0);
}

void BaseRS::selectAll()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("WholeStory"), 0);
}
