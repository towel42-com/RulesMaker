#include "OutlookObj.h"
#include <oaidl.h>
#include <QAxObject>

std::optional< Outlook::OlObjectClass > getObjectClass( QAxObject *item )
{
    if ( !item )
        return {};

    return static_cast< Outlook::OlObjectClass >( item->property( "Class" ).toInt() );
}

std::optional< Outlook::OlObjectClass > getObjectClass( IDispatch *item )
{
    if ( !item )
        return {};

    IDispatch *pdisp = (IDispatch *)nullptr;
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

    return {};
}

QString runStringMethod( OLECHAR *method, IDispatch *item )
{
    if ( !item )
        return false;
    IDispatch *pdisp = (IDispatch *)nullptr;
    DISPID dispid;
    auto result = item->GetIDsOfNames( IID_NULL, &method, 1, LOCALE_SYSTEM_DEFAULT, &dispid );
    if ( result == S_OK )
    {
        VARIANT resultant{};
        DISPPARAMS params{ 0 };
        EXCEPINFO excepInfo{};
        UINT argErr{ 0 };
        result = item->Invoke( dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &params, &resultant, &excepInfo, &argErr );
        if ( result == S_OK )
        {
            return QString::fromWCharArray( static_cast< wchar_t * >( resultant.bstrVal ) );
        }
    }
    return {};
}

bool hasMethod( OLECHAR *method, IDispatch *item )
{
    if ( !item )
        return false;
    IDispatch *pdisp = (IDispatch *)nullptr;
    DISPID dispid;
    auto result = item->GetIDsOfNames( IID_NULL, &method, 1, LOCALE_SYSTEM_DEFAULT, &dispid );
    return result == S_OK;
}

bool hasDelete( IDispatch *item )
{
    if ( !item )
        return false;

    return hasMethod( L"Delete", item );
}

bool deleteItem( IDispatch *item )
{
    if ( !hasDelete( item ) )
        return false;

    if ( !item )
        return {};

    IDispatch *pdisp = (IDispatch *)nullptr;
    DISPID dispid;
    OLECHAR *szMember = L"Delete";
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
            return true;
        }
    }
    return false;
}

QString getDescription( IDispatch *item )
{
    QString typeString;
    auto classType = getObjectClass( item );
    if ( classType.has_value() )
        typeString = toString( classType.value() );

    QString desc;
    if ( hasMethod( L"DisplayName", item ) )
    {
        desc = runStringMethod( L"DisplayName", item );
    }
    else if ( hasMethod( L"Name", item ) )
    {
        desc = runStringMethod( L"Name", item );
    }
    else if ( hasMethod( L"CurrentProfileName", item ) )
    {
        desc = runStringMethod( L"CurrentProfileName", item );
    }
    else if ( hasMethod( L"Subject", item ) )
    {
        desc = runStringMethod( L"Subject", item );
    }
    else if ( hasMethod( L"Caption", item ) )
    {
        desc = runStringMethod( L"Caption", item );
    }
    else if ( hasMethod( L"Formula", item ) )
    {
        desc = runStringMethod( L"Formula", item );
    }
    else if ( hasMethod( L"Filter", item ) )
    {
        desc = runStringMethod( L"Filter", item );
    }
    else if ( hasMethod( L"ConversationTopic", item ) )
    {
        desc = runStringMethod( L"ConversationTopic", item );
    }
    else
        return typeString;

    auto descText = ( QStringList() << typeString << desc ).join( " - " );
    return descText;
}