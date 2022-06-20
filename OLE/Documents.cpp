#include "StdAfx.h"
#include "Documents.h"
#include "Document.h"
#include "OLEMethod.h"

Documents::Documents(IDispatch* docs)
    : docs_(docs)
{

}

Documents::~Documents()
{
    docsList.clear();
    SafeRelease(docs_);
}

void Documents::closeAll()
{
    for (auto it = docsList.begin(); it != docsList.end(); ++it) {
        tDocumentSp sp = it->lock();
        if (sp)
            sp->close();
    }
    docsList.clear();
}


