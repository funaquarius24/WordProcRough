#include "StdAfx.h"
#include "Selection.h"
#include "OLEMethod.h"

/// ----------------------------------------------------------------------------
Selection::Selection( IDispatch* selection )
    : BaseRS(selection)
{

}

/// ----------------------------------------------------------------------------
int Selection::moveCursor( MoveDirection md, bool selectWhileMoving )
{
    VARIANT result;
    VariantInit(&result);

    VARIANT wdCharacter, wdExtend,Count;
    wdCharacter.vt   = VT_I4;
    wdCharacter.lVal = 1;
    Count.vt         = VT_I4;
    Count.lVal       = 1;
    wdExtend.vt      = VT_I4;
    wdExtend.lVal    = selectWhileMoving ? 1 : 0;

    switch (md) {
    case mdLeft:
        OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("MoveLeft"),  3, wdExtend, Count, wdCharacter);
        break;
    case mdRight:
        OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("MoveRight"), 3, wdExtend, Count, wdCharacter);
        break;
    case mdUp:
        wdCharacter.lVal = 5;
        OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("MoveUp"),    3, wdExtend, Count, wdCharacter);
        break;
    case mdDown:
        wdCharacter.lVal = 5;
        OLEMethod(DISPATCH_METHOD, &result, disp_, toLPOLESTR("MoveDown"),  3, wdExtend, Count, wdCharacter);
        break;
    }

    return result.intVal;
}

/// ----------------------------------------------------------------------------
// wstring_t Selection::getString( int length )
// {
//     return wstring_t();
// }

/// ----------------------------------------------------------------------------
void Selection::selectCurrentColor()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("SelectCurrentColor"), 0);
}

/// ----------------------------------------------------------------------------
void Selection::selectCurrentFont()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("SelectCurrentFont"), 0);
}

void Selection::copyFormat()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("CopyFormat"), 0);
}

void Selection::pasteFormat()
{
    OLEMethod(DISPATCH_METHOD, NULL, disp_, toLPOLESTR("PasteFormat"), 0);
}



