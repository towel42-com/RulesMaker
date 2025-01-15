#include "OutlookAPI.h"
#include "EmailAddress.h"

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

TEmailAddressList cleanResults( QStringList keys )
{
    TEmailAddressList retVal;
    keys = mergeStringLists( keys, {}, false );
    for ( auto &&ii : keys )
    {
        auto curr = CEmailAddress::fromKey( ii );
        if ( curr )
            retVal.push_back( curr );
    }

    return retVal;
}

TEmailAddressList cleanResults( TEmailAddressList values )
{
    QStringList keys;
    for ( auto &&ii : values )
    {
        keys << ii->key();
    }
    return cleanResults( keys );
}

TEmailAddressList mergeResults( TEmailAddressList lhs, const TEmailAddressList &rhs )
{
    QStringList keys;
    for ( auto &&ii : lhs )
    {
        keys << ii->key();
    }

    for ( auto &&ii : rhs )
    {
        keys << ii->key();
    }

    return cleanResults( keys );
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

TEmailAddressList COutlookAPI::getEmailAddresses( std::shared_ptr< Outlook::MailItem > &mailItem, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )   // returns the list of email addresses, display names
{
    if ( !mailItem )
        return {};
    return getEmailAddresses( mailItem.get(), types, contactTypes );
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::MailItem *mailItem, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )   // returns the list of email addresses, display names
{
    if ( !mailItem )
        return {};

    TEmailAddressList retVal;
    if ( isAddressType( types, EAddressTypes::eSender ) )
    {
        retVal = getEmailAddresses( mailItem->Sender(), types, contactTypes );
        bool isExchangeContact = isExchangeUser( mailItem->Sender() );
        if ( isExchangeContact )
            int xyz = 0;
        if ( isContactType( isExchangeContact, contactTypes ) )
        {
            auto emailAddress = mailItem->SenderEmailAddress();
            if ( isExchangeContact && !retVal.empty() )
                emailAddress = retVal.back()->emailAddress();
            retVal.push_back( std::make_shared< CEmailAddress >( emailAddress, mailItem->SentOnBehalfOfName(), isExchangeContact ) );
        }
    }

    auto curr = getEmailAddresses( mailItem->Recipients(), types, contactTypes );
    return mergeResults( retVal, curr );
}

//QStringList COutlookAPI::getEmailAddresses( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType, bool smtpOnly )
//{
//    return getEmailAddresses( getEmailAddresses( mailItem->Recipients(), getAddressTypes( recipientType, smtpOnly ) ) );
//}

QStringList COutlookAPI::getSenderEmailAddresses( Outlook::MailItem *mailItem )
{
    return CEmailAddress::getEmailAddresses( getEmailAddresses( mailItem, EAddressTypes::eSender ) );
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::Recipients *recipients, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )
{
    if ( !recipients )
        return {};

    TEmailAddressList retVal;

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
                useRecipient = isAddressType( types, EAddressTypes::eOriginator );
                break;
            case Outlook::OlMailRecipientType::olTo:
                useRecipient = isAddressType( types, EAddressTypes::eTo );
                break;
            case Outlook::OlMailRecipientType::olCC:
                useRecipient = isAddressType( types, EAddressTypes::eCC );
                break;
            case Outlook::OlMailRecipientType::olBCC:
                useRecipient = isAddressType( types, EAddressTypes::eBCC );
                break;
            default:
                break;
        }

        if ( !useRecipient )
            continue;

        auto curr = getEmailAddresses( recipient->AddressEntry(), types, contactTypes );
        retVal.insert( retVal.end(), curr.begin(), curr.end() );
    }
    return retVal;
}

//QStringList COutlookAPI::getEmailAddresses( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly )
//{
//    return getEmailAddresses( getEmailAddresses( recipients, getAddressTypes( recipientType, smtpOnly ) ) );
//}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::Recipient *recipient, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )
{
    if ( !recipient )
        return {};

    auto retVal = getEmailAddresses( recipient->AddressEntry(), types, contactTypes );
    if ( retVal.empty() )
    {
        bool isExchangeContact = isExchangeUser( recipient->AddressEntry() );
        if ( isExchangeContact )
            int xyz = 0;
        if ( isContactType( isExchangeContact, contactTypes ) )
            retVal.push_back( std::make_shared< CEmailAddress >( recipient->Address(), recipient->Name(), isExchangeContact ) );
    }

    return retVal;
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )
{
    if ( !addresses )
        return {};

    auto entries = addresses->AddressEntries();
    if ( !entries )
        return {};
    auto count = entries->Count();

    TEmailAddressList retVal;
    for ( int ii = 1; ii <= count; ++ii )
    {
        auto entry = entries->Item( ii );
        if ( !entry )
            continue;
        auto curr = getEmailAddresses( entry, types, contactTypes );
        retVal.insert( retVal.end(), curr.begin(), curr.end() );
    }

    return cleanResults( retVal );
}

//QStringList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, bool smtpOnly )
//{
//    return getEmailAddresses( getEmailAddresses( addresses, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ) );
//}

void COutlookAPI::displayEmail( const std::shared_ptr< Outlook::MailItem > &email ) const
{
    if ( email )
        email->Display();
}

bool COutlookAPI::isExchangeUser( Outlook::AddressEntry *address )
{
    return address && ( address->GetExchangeUser() || address->GetExchangeDistributionList() );
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )
{
    if ( !entries )
        return {};

    TEmailAddressList retVal;
    auto num = entries->Count();
    for ( int ii = 1; ii <= num; ++ii )
    {
        auto currItem = entries->Item( ii );
        if ( !currItem )
            continue;
        auto curr = getEmailAddresses( currItem, types, contactTypes );
        retVal.insert( retVal.end(), curr.begin(), curr.end() );
    }
    return cleanResults( retVal );
}

//QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, bool smtpOnly )
//{
//    return getEmailAddresses( getEmailAddresses( entries, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ) );
//}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )
{
    if ( !address )
        return {};
    auto type = address->AddressEntryUserType();

    TEmailAddressList retVal;

    if ( address->GetExchangeUser() )
    {
        if ( isContactType( contactTypes, EContactTypes::eOutlookContact ) )
        {
            retVal.push_back( std::make_shared< CEmailAddress >( address->GetExchangeUser()->PrimarySmtpAddress(), address->GetExchangeUser()->Name(), true ) );
        }
    }
    else if ( address->GetExchangeDistributionList() )
    {
        if ( isContactType( contactTypes, EContactTypes::eOutlookContact ) )
        {
            retVal.push_back( std::make_shared< CEmailAddress >( address->GetExchangeDistributionList()->PrimarySmtpAddress(), address->GetExchangeDistributionList()->Name(), true ) );
        }
    }
    else if ( address->GetContact() )
    {
        if ( isContactType( contactTypes, EContactTypes::eSMTPContact ) )
        {
            auto contact = address->GetContact();
            retVal.push_back( std::make_shared< CEmailAddress >( contact->Email1Address(), contact->Email1DisplayName(), false ) );   //
            retVal.push_back( std::make_shared< CEmailAddress >( contact->Email2Address(), contact->Email2DisplayName(), false ) );   //
            retVal.push_back( std::make_shared< CEmailAddress >( contact->Email3Address(), contact->Email3DisplayName(), false ) );
        }
    }
    else
    {
        if ( isContactType( contactTypes, EContactTypes::eSMTPContact ) )
            retVal.push_back( std::make_shared< CEmailAddress >( address->Address(), address->Name(), false ) );
    }

    return cleanResults( retVal );
}

//QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, bool smtpOnly /*= false*/ )
//{
//    return getEmailAddresses( getEmailAddresses( address, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ) );
//}
