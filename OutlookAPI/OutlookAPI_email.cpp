#include "OutlookAPI.h"
#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QWidget>

#include <QStringView>
#include <QMetaProperty>

#include <QVariant>
#include <iostream>
#include <oaidl.h>
#include "MSOUTL.h"
#include <QDebug>

COutlookAPI::TStringListPair cleanResults( COutlookAPI::TStringListPair values )
{
    values.first.removeAll( QString() );
    values.second.removeAll( QString() );
    values.first.removeDuplicates();
    values.second.removeDuplicates();
    return values;
}

COutlookAPI::TStringListPair mergeResults( COutlookAPI::TStringListPair lhs, const COutlookAPI::TStringListPair &rhs )
{
    lhs.first << rhs.first;
    lhs.second << rhs.second;
    return cleanResults( lhs );
}

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

std::shared_ptr< Outlook::MailItem > COutlookAPI::getEmailItem( const std::shared_ptr< Outlook::Items > &items, int num )
{
    if ( !items || !num || ( num > items->Count() ) )
        return {};

    auto item = items->Item( num );
    if ( !item )
        return {};

    if ( getObjectClass( item ) == Outlook::OlObjectClass::olMail )
    {
        return getEmailItem( item );
    }
    return {};
}

std::shared_ptr< Outlook::MailItem > COutlookAPI::getEmailItem( IDispatch *item )
{
    if ( !item )
        return {};
    return connectToException( std::make_shared< Outlook::MailItem >( item ) );
}

COutlookAPI::TStringListPair COutlookAPI::getEmailAddresses( std::shared_ptr< Outlook::MailItem > &mailItem, EAddressTypes types )   // returns the list of email addresses, display names
{
    if ( !mailItem )
        return {};
    return getEmailAddresses( mailItem.get(), types );
}

COutlookAPI::TStringListPair COutlookAPI::getEmailAddresses( Outlook::MailItem *mailItem, EAddressTypes types )   // returns the list of email addresses, display names
{
    if ( !mailItem )
        return {};

    TStringListPair retVal;
    if ( ( types & EAddressTypes::eSender ) != 0 )
    {
        retVal = getEmailAddresses( mailItem->Sender(), types );
    }

    auto curr = getEmailAddresses( mailItem->Recipients(), types );
    return mergeResults( retVal, curr );
}


QStringList COutlookAPI::getEmailAddresses( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType, bool smtpOnly )
{
    return getEmailAddresses( mailItem->Recipients(), getAddressTypes( recipientType, smtpOnly ) ).first;
}

QStringList COutlookAPI::getSenderEmailAddresses( Outlook::MailItem *mailItem )
{
    return getEmailAddresses( mailItem, EAddressTypes::eSender ).first;
}

COutlookAPI::TStringListPair COutlookAPI::getEmailAddresses( Outlook::Recipients *recipients, EAddressTypes types )
{
    if ( !recipients )
        return {};

    TStringListPair retVal;

    auto numRecipients = recipients->Count();
    for ( int ii = 1; ii <= numRecipients; ++ii )
    {
        auto recipient = recipients->Item( ii );
        if ( !recipient )
            continue;
        bool useRecipient = false;
        auto recipientType = static_cast< Outlook::OlMailRecipientType >( recipient->Type() );
        switch ( recipientType )
        {
            case Outlook::OlMailRecipientType::olOriginator:
                useRecipient = ( types & EAddressTypes::eOriginator ) != 0;
                break;
            case Outlook::OlMailRecipientType::olTo:
                useRecipient = ( types & EAddressTypes::eTo ) != 0;
                break;
            case Outlook::OlMailRecipientType::olCC:
                useRecipient = ( types & EAddressTypes::eCC ) != 0;
                break;
            case Outlook::OlMailRecipientType::olBCC:
                useRecipient = ( types & EAddressTypes::eBCC ) != 0;
                break;
            default:
                break;
        }

        if ( !useRecipient )
            continue;

        auto curr = getEmailAddresses( recipient->AddressEntry(), types );
        retVal.first << curr.first;
        retVal.second << curr.second;
    }
    return retVal;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly )
{
    return getEmailAddresses( recipients, getAddressTypes( recipientType, smtpOnly ) ).first;
}

COutlookAPI::TStringListPair COutlookAPI::getEmailAddresses( Outlook::Recipient *recipient, EAddressTypes types )
{
    if ( !recipient )
        return {};

    auto retVal = getEmailAddresses( recipient->AddressEntry(), types );
    if ( retVal.first.isEmpty() )
        retVal.first << recipient->Address();
    if ( retVal.second.isEmpty() )
        retVal.first << recipient->Name();

    return retVal;
}

COutlookAPI::TStringListPair COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, EAddressTypes types )
{
    if ( !addresses )
        return {};

    auto entries = addresses->AddressEntries();
    if ( !entries )
        return {};
    auto count = entries->Count();

    TStringListPair retVal;
    for ( int ii = 1; ii <= count; ++ii )
    {
        auto entry = entries->Item( ii );
        if ( !entry )
            continue;
        auto currEmails = getEmailAddresses( entry, types );
        retVal.first << currEmails.first;
        retVal.second << currEmails.second;
    }

    return cleanResults( retVal );
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, bool smtpOnly )
{
    return getEmailAddresses( addresses, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ).first;
}

void COutlookAPI::displayEmail( const std::shared_ptr< Outlook::MailItem > &email ) const
{
    if ( email )
        email->Display();
}

COutlookAPI::TStringListPair COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, EAddressTypes types )
{
    if ( !entries )
        return {};

    TStringListPair retVal;
    auto num = entries->Count();
    for ( int ii = 1; ii <= num; ++ii )
    {
        auto currItem = entries->Item( ii );
        if ( !currItem )
            continue;
        auto currEmails = getEmailAddresses( currItem, types );
        retVal.first << currEmails.first;
        retVal.second << currEmails.second;
    }
    return cleanResults( retVal );
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, bool smtpOnly )
{
    return getEmailAddresses( entries, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ).first;
}

COutlookAPI::TStringListPair COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, EAddressTypes types )
{
    if ( !address )
        return {};
    auto type = address->AddressEntryUserType();

    bool smtpOnly = ( types & COutlookAPI::EAddressTypes::eSMTPOnly ) != 0;

    TStringListPair retVal;

    if ( address->GetExchangeUser() )
    {
        if ( !smtpOnly )
        {
            retVal.first << address->GetExchangeUser()->PrimarySmtpAddress();
            retVal.second << address->GetExchangeUser()->Name();
        }
    }
    else if ( address->GetExchangeDistributionList() )
    {
        if ( smtpOnly )
        {
            retVal.first << address->GetExchangeDistributionList()->PrimarySmtpAddress();
            retVal.second << address->GetExchangeDistributionList()->Name();
        }
    }
    else if ( address->GetContact() )
    {
        auto contact = address->GetContact();
        retVal.first << contact->Email1Address() << contact->Email2Address() << contact->Email3Address();
        retVal.second << contact->Email1DisplayName() << contact->Email2DisplayName() << contact->Email3DisplayName();
    }
    else
    {
        retVal.first << address->Address();
        retVal.second << address->Name();
    }

    return cleanResults( retVal );
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, bool smtpOnly /*= false*/ )
{
    return getEmailAddresses( address, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ).first;
}
