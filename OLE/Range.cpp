#include "StdAfx.h"
#include "Range.h"
#include "OLEMethod.h"
#include "Selection.h"


tRangeSp Note::getRange() const
{
    return tRangeSp(new Range(getPropertyDispatch(disp_, toLPOLESTR("Range"))) );
}


tNoteSp Notes::getItem( int index )
{
    VARIANT result;
    VariantInit(&result);

    VARIANT x;
    VariantInit(&x);
    x.vt = VT_I4;
    x.lVal = index;
    OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("Item"), 1, x);

    return tNoteSp(new Note(result.pdispVal));
}

int Notes::getCount() const
{
    return getPropertyInt(disp_, toLPOLESTR("Count"));
}


Range::Range( IDispatch* range )
    : BaseRS(range)
{
}

tRangeSp Range::getNextStoryRange()
{
    IDispatch* disp = getPropertyDispatch(disp_, toLPOLESTR("NextStoryRange"));
    if (disp)
        return tRangeSp(new Range(disp) );
    return tRangeSp();
}

tRangeSp Range::getNext(int wdUnit, int count)
{
    VARIANT result;
    VariantInit(&result);
    
    VARIANT unt;
    VariantInit(&unt);
    unt.vt = VT_I4;
    unt.intVal = wdUnit;

    OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("Next"), 1, unt);

    if (result.pdispVal)
        return tRangeSp(new Range(result.pdispVal) );
    return tRangeSp();
}

void Range::autoFormat()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("AutoFormat"), 0);
}

int Range::getHighlightColorIndex()
{
    return getPropertyInt(disp_, toLPOLESTR("HighlightColorIndex"));
}

tFootnotesSp Range::getFootnotes()
{
    return tFootnotesSp(new Footnotes(getPropertyDispatch(disp_, toLPOLESTR("Footnotes"))) );
}

tFindSp Range::getFind()
{
    return tFindSp(new Find(getPropertyDispatch(disp_, toLPOLESTR("Find"))) );
}

void Range::collapse()
{
    VARIANT result;
    VariantInit(&result);

    VARIANT x;
    VariantInit(&x);
    x.vt = VT_I4;
    x.lVal = 0;
    OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("Collapse"), 1, x);
}

void Range::setRange( int startPos, int endPos )
{
    VARIANT result;
    VariantInit(&result);

    VARIANT x, y;
    VariantInit(&x);
    x.vt = VT_I4;
    x.intVal = startPos;

    VariantInit(&y);
    y.vt = VT_I4;
    y.intVal = endPos;

    OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("SetRange"), 2, y, x);
}

tTextRetrievalModeSp Range::textRetrievalMode()
{
    return tTextRetrievalModeSp(new TextRetrievalMode(getPropertyDispatch(disp_, toLPOLESTR("TextRetrievalMode"))) );
}

tRangeSp Range::duplicate()
{
    return tRangeSp(new Range(getPropertyDispatch(disp_, toLPOLESTR("Duplicate"))) );
}

tParagraphFormatSp Range::getParagraphFormat()
{
    return tParagraphFormatSp(new ParagraphFormat(getPropertyDispatch(disp_, toLPOLESTR("ParagraphFormat"))) );
}

// 
// tBaseObjectSp Range::getStyle()
// {
//     return tBaseObjectSp(new BaseObject(getPropertyDispatch(disp_, toLPOLESTR("Style"))) );
// }
// 
// void Range::setStyle( const tBaseObjectSp& obj )
// {
//     setPropertyDispatch(disp_, obj->getIDispatch(), toLPOLESTR("Style"));
// }

StoryRanges::StoryRanges( IDispatch* range )
    : range_(range)
{

}

StoryRanges::~StoryRanges()
{
    SafeRelease(range_);
}

int StoryRanges::getCount()
{
    return getPropertyInt(range_, toLPOLESTR("Count"));
}


Characters::Characters( IDispatch* chars )
    : chars_(chars)
{

}

Characters::~Characters()
{
    SafeRelease(chars_);
}

int Characters::getCount()
{
    return getPropertyInt(chars_, toLPOLESTR("Count"));
}

tRangeSp Characters::getItem( int index )
{
    VARIANT result;
    VARIANT input;
    VariantInit(&result);
    VariantInit(&input);

    input.vt = VT_I4;
    input.intVal = index;
    OLEMethod(DISPATCH_PROPERTYGET, &result, chars_, toLPOLESTR("Item"), 1, input);
    
    return tRangeSp(new Range(result.pdispVal));
}


tRangeSp Characters::getFirst()
{
    return tRangeSp(new Range(getPropertyDispatch(chars_, toLPOLESTR("First"))) );
}


Paragraph::Paragraph( IDispatch* paragraph )
    : paragraph_(paragraph)
{

}

Paragraph::~Paragraph()
{
    SafeRelease(paragraph_);
}

tParagraphSp Paragraph::getNext()
{
    VARIANT result;
    VariantInit(&result);
    OLEMethod(DISPATCH_METHOD, &result, paragraph_, toLPOLESTR("Next"), 0);

    if (result.pdispVal)
        return tParagraphSp(new Paragraph(result.pdispVal) );
    return tParagraphSp();
}

IDispatch* Paragraph::getIDispatch() const
{
    return paragraph_;
}

tRangeSp Paragraph::getRange()
{
    return tRangeSp(new Range(getPropertyDispatch(paragraph_, toLPOLESTR("Range"))) );
}

tBaseObjectSp Paragraph::getStyle()
{
    return tBaseObjectSp(new BaseObject(getPropertyDispatch(paragraph_, toLPOLESTR("Style"))) );
}

tBaseObjectSp Paragraph::getFormat()
{
    return tBaseObjectSp(new BaseObject(getPropertyDispatch(paragraph_, toLPOLESTR("Format"))) );
}

void Paragraph::setStyle( const tBaseObjectSp& obj )
{
    setPropertyDispatch(paragraph_, obj->getIDispatch(), toLPOLESTR("Style"));
}

void Paragraph::setFormat( const tBaseObjectSp& obj )
{
    setPropertyDispatch(paragraph_, obj->getIDispatch(), toLPOLESTR("Format"));
}

Paragraphs::Paragraphs( IDispatch* paragraph )
    : paragraphs_(paragraph)
{
}

Paragraphs::~Paragraphs()
{
    SafeRelease(paragraphs_);
}

int Paragraphs::getCount()
{
    return getPropertyInt(paragraphs_, toLPOLESTR("Count"));
}

tParagraphSp Paragraphs::getItem( int index )
{
    VARIANT result;
    VariantInit(&result);

    VARIANT x;
    //VariantInit(&input);
    x.vt = VT_I4;
    x.intVal = index;
    OLEMethod(DISPATCH_METHOD, &result, paragraphs_, toLPOLESTR("Item"), 1, x);

    return tParagraphSp(new Paragraph(result.pdispVal));
}

tParagraphSp Paragraphs::getFirst()
{
    return tParagraphSp(new Paragraph(getPropertyDispatch(paragraphs_, toLPOLESTR("First"))) );
}

tStyleSp Styles::getItem( int index )
{
    VARIANT result;
    VariantInit(&result);

    VARIANT x;
    //VariantInit(&input);
    x.vt = VT_I4;
    x.intVal = index;
    OLEMethod(DISPATCH_METHOD, &result, styles_, toLPOLESTR("Item"), 1, x);

    return tStyleSp(new Style(result.pdispVal));
}

int Styles::getCount()
{
    return getPropertyInt(styles_, toLPOLESTR("Count"));
}



Style::Style( IDispatch* style )
    : style_(style)
{

}

Style::~Style()
{
    SafeRelease(style_);
}


/// ----------------------------------------------------------------------------
IDispatch* Collection::getItem( int index )
{
    VARIANT result;
    VariantInit(&result);

    VARIANT x;
    x.vt = VT_I4;
    x.intVal = index;
    OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("Item"), 1, x);

    return result.pdispVal;
}

int Collection::getCount()
{
    return getPropertyInt(disp_, toLPOLESTR("Count"));
}

tHeadersFootersSp Section::getFooters()
{
    return tHeadersFootersSp(new HeadersFooters(getPropertyDispatch(disp_, toLPOLESTR("Footers"))) );
}

tHeadersFootersSp Section::getHeaders()
{
    return tHeadersFootersSp(new HeadersFooters(getPropertyDispatch(disp_, toLPOLESTR("Headers"))) );
}

tRangeSp HeaderFooter::getRange()
{
    return tRangeSp(new Range(getPropertyDispatch(disp_, toLPOLESTR("Range"))) );
}

void TextRetrievalMode::setFieldCodes( short value )
{
    setPropertyBool(disp_, toLPOLESTR("IncludeFieldCodes"), value);
}

short TextRetrievalMode::getFieldCodes() const
{
    return getPropertyBool(disp_, toLPOLESTR("IncludeFieldCodes"));
}

void TextRetrievalMode::setHiddenText( short value )
{
    setPropertyBool(disp_, toLPOLESTR("IncludeHiddenText"), value);
}

short TextRetrievalMode::getHiddenText() const
{
    return getPropertyBool(disp_, toLPOLESTR("IncludeHiddenText"));
}

void ParagraphFormat::setAlignment( int value )
{
    setPropertyInt(disp_, toLPOLESTR("Alignment"), value);
}

int ParagraphFormat::getAlignment() const
{
    return getPropertyInt(disp_, toLPOLESTR("Alignment"));
}

void ParagraphFormat::setLineSpacing( float value )
{
    setPropertyFloat(disp_, toLPOLESTR("LineSpacing"), value);
}

float ParagraphFormat::getLineSpacing() const
{
    return getPropertyFloat(disp_, toLPOLESTR("LineSpacing"));
}

void ParagraphFormat::reset()
{
    OLEMethod(DISPATCH_METHOD, 0, disp_, toLPOLESTR("Reset"), 0);
}
