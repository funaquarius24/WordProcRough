#include "StdAfx.h"
#include "Document.h"
#include "OLEMethod.h"

Document::Document(IDispatch* doc)
    : doc_(doc)
{
}

Document::~Document()
{
    SafeRelease(doc_);
}

void Document::close()
{
    if (doc_) {
        OLEMethod(DISPATCH_METHOD, NULL, doc_, toLPOLESTR("Close"), 0);
        SafeRelease(doc_);
    }
}

void Document::save()
{
    if (doc_) 
        OLEMethod(DISPATCH_METHOD, NULL, doc_, toLPOLESTR("Save"), 0);
}


tRangeSp Document::getContent()
{
    return tRangeSp(new Range(getPropertyDispatch(doc_, toLPOLESTR("Content"))) );
}

tStoryRangesSp Document::getStoryRanges()
{
    return tStoryRangesSp(new StoryRanges(getPropertyDispatch(doc_, toLPOLESTR("StoryRanges"))) );
}

tCharactersSp Document::getCharacters()
{
    return tCharactersSp(new Characters(getPropertyDispatch(doc_, toLPOLESTR("Characters"))) );
}

tParagraphsSp Document::getParagraphs()
{
    return tParagraphsSp(new Paragraphs(getPropertyDispatch(doc_, toLPOLESTR("Paragraphs"))) );
}

tFootnotesSp Document::getFootnotes()
{
    return tFootnotesSp(new Footnotes(getPropertyDispatch(doc_, toLPOLESTR("Footnotes"))) );
}

tStylesSp Document::getStyles()
{
    return tStylesSp(new Styles(getPropertyDispatch(doc_, toLPOLESTR("Styles"))) );
}

tFormFieldsSp Document::getFormFields()
{
    return tFormFieldsSp(new FormFields(getPropertyDispatch(doc_, toLPOLESTR("FormFields"))) );
}

tSectionsSp Document::getSections()
{
    return tSectionsSp(new Sections(getPropertyDispatch(doc_, toLPOLESTR("Sections"))) );
}

tSentencesSp Document::getSentences()
{
    return tSentencesSp(new Sentences(getPropertyDispatch(doc_, toLPOLESTR("Sentences"))) );
}
