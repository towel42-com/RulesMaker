#include "OutlookAPI.h"
#include "EmailAddress.h"

#include "MSOUTL.h"
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

std::pair< bool, QString > COutlookAPI::canDeleteItem( IDispatch *item )
{
    std::pair< bool, QString > retVal{ false, QString() };
    if ( !item )
        return retVal;

    item->AddRef();
    auto axObject = QAxObject( (IUnknown *)item, nullptr );
    auto properties = axObject.propertyBag();

    auto desc = Outlook::toString( static_cast< Outlook::OlObjectClass >( axObject.dynamicCall( "Class()" ).toInt() ) );
    auto metaObject = axObject.metaObject();
    if ( !metaObject )
    {
        return retVal;
    }

    auto methodIndex = metaObject->indexOfMethod( "Delete()" );
    return { methodIndex != -1, desc };
}

bool COutlookAPI::deleteItem( IDispatch *item )
{
    if ( !canDeleteItem( item ).first )
        return false;

    auto axObject = QAxObject( (IUnknown *)item, nullptr );

    axObject.dynamicCall( "Delete()" );
    return true;
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
