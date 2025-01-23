#include "OutlookObj.h"
#include <oaidl.h>
#include <QAxObject>

Outlook::OlObjectClass getObjectClass( QAxObject *item )
{
    auto retVal = static_cast< Outlook::OlObjectClass >( -1 );;
    if ( item )
        retVal = static_cast< Outlook::OlObjectClass >( item->property( "Class" ).toInt() );
    return retVal;
}

Outlook::OlObjectClass getObjectClass( IDispatch *item )
{
    if ( !item )
        return {};

    IDispatch *pdisp = (IDispatch *)NULL;
    DISPID dispid;
    OLECHAR *szMember = L"Class";
    auto result = item->GetIDsOfNames( IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid );

    if ( result == S_OK )
    {
        VARIANT resultant{};
        DISPPARAMS params{ 0 };
        EXCEPINFO excepInfo{};
        UINT argErr{ 0 };
        result = item->Invoke( dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &params, &resultant, &excepInfo, &argErr );
        if ( result == S_OK )
        {
            return static_cast< Outlook::OlObjectClass >( resultant.lVal );
        }
    }

    return getObjectClass( &QAxObject( item ) );
}