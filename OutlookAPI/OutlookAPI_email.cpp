#include "OutlookAPI.h"
#include "EmailAddress.h"

#include "OutlookLib/MSOUTL.h"
#include <oaidl.h>

std::pair< std::shared_ptr< Outlook::Items >, int > COutlookAPI::getEmailItemsForRootFolder()
{
    auto folder = COutlookAPI::instance()->rootFolder();
    if ( !folder )
        return {};
    auto fn = folder->FullFolderPath();

    std::shared_ptr< Outlook::Items > retVal;

    auto items = folder->Items();
    if ( items )
    {
        auto limitToUnread = onlyProcessUnread();
        if ( limitToUnread && ( items->Count() < 200 ) )
            limitToUnread = !processAllEmailWhenLessThan200Emails();

        if ( limitToUnread )
        {
            auto subItems = items->Restrict( "[UNREAD]=TRUE" );
            if ( subItems )
                retVal = getItems( subItems );
        }

        if ( !retVal )
            retVal = COutlookAPI::instance()->getItems( items );
    }
    if ( !retVal )
        return {};
    return { retVal, retVal->Count() };
}

IDispatch *COutlookAPI::getItem( const std::shared_ptr< Outlook::Items > &items, int num )
{
    if ( !items || !num || ( num > items->Count() ) )
        return nullptr;

    auto item = items->Item( num );
    if ( !item )
        return item;

    return item;
}

std::shared_ptr< Outlook::MailItem > COutlookAPI::getEmailItem( IDispatch *item )
{
    if ( !item || ( getObjectClass( item ) != Outlook::OlObjectClass::olMail ) )
        return {};
    return connectToException( std::make_shared< Outlook::MailItem >( item ) );
}

std::tuple< bool, QString, std::unique_ptr< QAxObject > > COutlookAPI::canDeleteItem( IDispatch *item )
{
    std::tuple< bool, QString, std::unique_ptr< QAxObject > > retVal{ false, QString(), nullptr };
    if ( !item )
        return retVal;

    item->AddRef();
    auto axObject = std::make_unique< QAxObject >( (IUnknown *)item, nullptr );
    auto properties = axObject->propertyBag();
    auto metaObject = axObject->metaObject();
    if ( !metaObject )
    {
        return retVal;
    }

    auto methodIndex = metaObject->indexOfMethod( "Delete()" );
    auto desc = Outlook::toString( static_cast< Outlook::OlObjectClass >( axObject->dynamicCall( "Class()" ).toInt() ) );
    return { methodIndex != -1, desc, std::move( axObject ) };
}

std::pair< bool, QString > COutlookAPI::deleteItem( IDispatch *item )
{
    auto &&[ canDelete, desc, axObj ] = canDeleteItem( item );

    if ( canDelete )
        axObj->dynamicCall( "Delete()" );

    return { canDelete, desc };
}

void COutlookAPI::displayEmail( const std::shared_ptr< Outlook::MailItem > &email ) const
{
    if ( email )
        email->Display();
}

bool COutlookAPI::isExchangeUser( Outlook::AddressEntry *address )
{
    return address && ( address->GetExchangeUser() || address->GetExchangeDistributionList() );
}
