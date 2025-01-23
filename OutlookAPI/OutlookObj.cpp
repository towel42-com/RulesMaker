#include "OutlookObj.h"
#include <oaidl.h>
#include <QAxObject>

std::optional< Outlook::OlObjectClass > getObjectClass( QAxObject *item )
{
    if ( !item )
        return {};

    return static_cast< Outlook::OlObjectClass >( item->property( "Class" ).toInt() );
}

struct SOleMethod
{
    SOleMethod( OLECHAR *method, IDispatch *item ) :
        fMethod( method ),
        fItem( item )
    {
        if ( !fItem )
            return;

        auto result = item->GetIDsOfNames( IID_NULL, &fMethod, 1, LOCALE_SYSTEM_DEFAULT, &fDispid );
        if ( !result == S_OK )
            fDispid = -1;
    }

    bool hasMethod() const { return fItem && ( fDispid != -1 ); }

    QString runStringMethod()
    {
        auto resultant = run();
        if ( !resultant.has_value() )
            return {};

        return QString::fromWCharArray( resultant.value().bstrVal );
    };

    bool runBoolMethod()
    {
        auto resultant = run();
        return resultant.has_value();
    };

    LONG runLongMethod()
    {
        auto resultant = run();
        if ( !resultant.has_value() )
            return -1;

        return resultant.value().lVal;
    }
private:
    OLECHAR *fMethod{ nullptr };
    IDispatch *fItem{ nullptr };
    DISPID fDispid{ -1 };

    std::optional< VARIANT > run()
    {
        if ( !hasMethod() )
            return {};
        VARIANT resultant{};
        DISPPARAMS params{ 0 };
        EXCEPINFO excepInfo{};
        UINT argErr{ 0 };
        auto result = fItem->Invoke( fDispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &params, &resultant, &excepInfo, &argErr );
        if ( result == S_OK )
        {
            return resultant;
        }
        return {};
    }
};

bool hasMethod( OLECHAR *method, IDispatch *item )
{
    if ( !item )
        return false;

    DISPID dispid;
    auto result = item->GetIDsOfNames( IID_NULL, &method, 1, LOCALE_SYSTEM_DEFAULT, &dispid );
    return result == S_OK;
}

bool hasDelete( IDispatch *item )
{
    SOleMethod method( L"Delete", item );
    return method.hasMethod();
}

bool deleteItem( IDispatch *item )
{
    SOleMethod method( L"Delete", item );
    return method.runBoolMethod();
}

QString getDescription( IDispatch *item )
{
    QString typeString;
    auto classType = getObjectClass( item );
    if ( classType.has_value() )
        typeString = toString( classType.value() );

    QString desc;
    SOleMethod displayNameMethod( L"DisplayName", item );
    SOleMethod nameMethod( L"Name", item );
    SOleMethod currentProfileNameMethod( L"CurrentProfileName", item );
    SOleMethod subjectMethod( L"Subject", item );
    SOleMethod captionMethod( L"Caption", item );
    SOleMethod formulaMethod( L"Formula", item );
    SOleMethod filterMethod( L"Filter", item );
    SOleMethod conversationTopicMethod( L"ConversationTopic", item );

    if ( displayNameMethod.hasMethod() )
    {
        desc = displayNameMethod.runStringMethod();
    }
    else if ( nameMethod.hasMethod() )
    {
        desc = nameMethod.runStringMethod();
    }
    else if ( currentProfileNameMethod.hasMethod() )
    {
        desc = currentProfileNameMethod.runStringMethod();
    }
    else if ( subjectMethod.hasMethod() )
    {
        desc = subjectMethod.runStringMethod();
    }
    else if ( captionMethod.hasMethod() )
    {
        desc = captionMethod.runStringMethod();
    }
    else if ( formulaMethod.hasMethod() )
    {
        desc = formulaMethod.runStringMethod();
    }
    else if ( filterMethod.hasMethod() )
    {
        desc = filterMethod.runStringMethod();
    }
    else if ( conversationTopicMethod.hasMethod() )
    {
        desc = conversationTopicMethod.runStringMethod();
    }
    else
        return typeString;

    auto descText = ( QStringList() << typeString << desc ).join( " - " );
    return descText;
}

std::optional< Outlook::OlObjectClass > getObjectClass( IDispatch *item )
{
    SOleMethod method( L"Class", item );
    if ( !method.hasMethod() )
        return {};
    return static_cast< Outlook::OlObjectClass >( method.runLongMethod() );
}