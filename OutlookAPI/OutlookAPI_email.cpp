#include "OutlookAPI.h"
#include "EmailAddress.h"

#include "MSOUTL.h"

std::pair< COutlookObj< Outlook::_Items >, int > COutlookAPI::getEmailItemsForRootFolder()
{
    auto folder = COutlookAPI::instance()->rootFolder();
    if ( !folder )
        return {};
    auto fn = folder->FullFolderPath();

    COutlookObj< Outlook::_Items > retVal;

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

COutlookObj< Outlook::MailItem > COutlookAPI::getEmailItem( const COutlookObj< Outlook::_Items > &items, int num )
{
    if ( !items || !num || ( num > items->Count() ) )
        return {};

    auto item = items->Item( num );
    if ( !item )
        return {};

    COutlookObj< Outlook::MailItem > mailItem( item );
    if ( mailItem )
        return mailItem;
    return {};
}

COutlookObj< Outlook::MailItem > COutlookAPI::getEmailItem( IDispatch *item )
{
    if ( !item )
        return {};
    return COutlookObj< Outlook::MailItem >( item );
}

void COutlookAPI::displayEmail( const COutlookObj< Outlook::MailItem > &email ) const
{
    if ( email )
        email->Display();
}

bool COutlookAPI::isExchangeUser( Outlook::AddressEntry *address )
{
    return address && ( address->GetExchangeUser() || address->GetExchangeDistributionList() );
}

