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

COutlookAPI::TStringPairList cleanResults( QStringList keys )
{
    COutlookAPI::TStringPairList retVal;
    keys = mergeStringLists( keys, {}, false );
    for ( auto &&ii : keys )
    {
        auto split = ii.split( "<<<BREAK>>>" );
        if ( split.size() != 2 )
            continue;
        retVal << std::make_pair( split.first(), split.last() );
    }

    return retVal;
}

COutlookAPI::TStringPairList cleanResults( COutlookAPI::TStringPairList values )
{
    QStringList keys;
    for ( auto &&ii : values )
    {
        keys << ii.first + "<<<BREAK>>>" + ii.second;
    }
    return cleanResults( keys );
}

COutlookAPI::TStringPairList mergeResults( COutlookAPI::TStringPairList lhs, const COutlookAPI::TStringPairList &rhs )
{
    QStringList keys;
    for ( auto &&ii : lhs )
    {
        keys << ii.first + "<<<BREAK>>>" + ii.second;
    }

    for ( auto &&ii : rhs )
    {
        keys << ii.first + "<<<BREAK>>>" + ii.second;
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

COutlookAPI::TStringPairList COutlookAPI::getEmailAddresses( std::shared_ptr< Outlook::MailItem > &mailItem, EAddressTypes types )   // returns the list of email addresses, display names
{
    if ( !mailItem )
        return {};
    return getEmailAddresses( mailItem.get(), types );
}

COutlookAPI::TStringPairList COutlookAPI::getEmailAddresses( Outlook::MailItem *mailItem, EAddressTypes types )   // returns the list of email addresses, display names
{
    if ( !mailItem )
        return {};

    TStringPairList retVal;
    if ( ( types & EAddressTypes::eSender ) != 0 )
    {
        retVal = getEmailAddresses( mailItem->Sender(), types );
        bool smtpOnly = ( types & COutlookAPI::EAddressTypes::eSMTPOnly ) != 0;
        if ( smtpOnly && !isExchangeUser( mailItem->Sender() ) )
            retVal << std::make_pair( mailItem->SenderEmailAddress(), mailItem->SentOnBehalfOfName() );
    }

    auto curr = getEmailAddresses( mailItem->Recipients(), types );
    return mergeResults( retVal, curr );
}

QStringList COutlookAPI::getEmailAddresses( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType, bool smtpOnly )
{
    return getEmailAddresses( getEmailAddresses( mailItem->Recipients(), getAddressTypes( recipientType, smtpOnly ) ) );
}

QStringList COutlookAPI::getSenderEmailAddresses( Outlook::MailItem *mailItem )
{
    return getEmailAddresses( getEmailAddresses( mailItem, EAddressTypes::eSender ) );
}

COutlookAPI::TStringPairList COutlookAPI::getEmailAddresses( Outlook::Recipients *recipients, EAddressTypes types )
{
    if ( !recipients )
        return {};

    TStringPairList retVal;

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
        retVal << curr;
    }
    return retVal;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly )
{
    return getEmailAddresses( getEmailAddresses( recipients, getAddressTypes( recipientType, smtpOnly ) ) );
}

COutlookAPI::TStringPairList COutlookAPI::getEmailAddresses( Outlook::Recipient *recipient, EAddressTypes types )
{
    if ( !recipient )
        return {};

    auto retVal = getEmailAddresses( recipient->AddressEntry(), types );
    if ( retVal.empty() )
    {
        retVal << std::make_pair( recipient->Address(), recipient->Name() );
    }

    return retVal;
}

COutlookAPI::TStringPairList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, EAddressTypes types )
{
    if ( !addresses )
        return {};

    auto entries = addresses->AddressEntries();
    if ( !entries )
        return {};
    auto count = entries->Count();

    TStringPairList retVal;
    for ( int ii = 1; ii <= count; ++ii )
    {
        auto entry = entries->Item( ii );
        if ( !entry )
            continue;
        auto currEmails = getEmailAddresses( entry, types );
        retVal << currEmails;
    }

    return cleanResults( retVal );
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, bool smtpOnly )
{
    return getEmailAddresses( getEmailAddresses( addresses, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ) );
}

void COutlookAPI::displayEmail( const std::shared_ptr< Outlook::MailItem > &email ) const
{
    if ( email )
        email->Display();
}

QStringList COutlookAPI::getEmailAddresses( const TStringPairList &emailAddresses )
{
    QStringList retVal;
    for ( auto &&ii : emailAddresses )
    {
        retVal << ii.first;
    }
    return retVal;
}

QStringList COutlookAPI::getDisplayNames( const TStringPairList &emailAddresses )
{
    QStringList retVal;
    for ( auto &&ii : emailAddresses )
    {
        retVal << ii.second;
    }
    return retVal;
}

bool COutlookAPI::isExchangeUser( Outlook::AddressEntry *address )
{
    return address && ( address->GetExchangeUser() || address->GetExchangeDistributionList() );
}

COutlookAPI::TStringPairList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, EAddressTypes types )
{
    if ( !entries )
        return {};

    TStringPairList retVal;
    auto num = entries->Count();
    for ( int ii = 1; ii <= num; ++ii )
    {
        auto currItem = entries->Item( ii );
        if ( !currItem )
            continue;
        auto currEmails = getEmailAddresses( currItem, types );
        retVal << currEmails;
    }
    return cleanResults( retVal );
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, bool smtpOnly )
{
    return getEmailAddresses( getEmailAddresses( entries, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ) );
}

COutlookAPI::TStringPairList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, EAddressTypes types )
{
    if ( !address )
        return {};
    auto type = address->AddressEntryUserType();

    bool smtpOnly = ( types & COutlookAPI::EAddressTypes::eSMTPOnly ) != 0;

    TStringPairList retVal;

    if ( address->GetExchangeUser() )
    {
        if ( !smtpOnly )
        {
            retVal << std::make_pair( address->GetExchangeUser()->PrimarySmtpAddress(), address->GetExchangeUser()->Name() );
        }
    }
    else if ( address->GetExchangeDistributionList() )
    {
        if ( smtpOnly )
        {
            retVal << std::make_pair( address->GetExchangeDistributionList()->PrimarySmtpAddress(), address->GetExchangeDistributionList()->Name() );
        }
    }
    else if ( address->GetContact() )
    {
        auto contact = address->GetContact();
        retVal   //
            << std::make_pair( contact->Email1Address(), contact->Email1DisplayName() )   //
            << std::make_pair( contact->Email2Address(), contact->Email2DisplayName() )   //
            << std::make_pair( contact->Email3Address(), contact->Email3DisplayName() );
    }
    else
    {
        retVal << std::make_pair( address->Address(), address->Name() );
    }

    return cleanResults( retVal );
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, bool smtpOnly /*= false*/ )
{
    return getEmailAddresses( getEmailAddresses( address, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ) );
}
