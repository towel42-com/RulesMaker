#include "OutlookAPI.h"
#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QWidget>

#include <QStringView>
#include <QMetaProperty>
#include <QSettings>

#include <QVariant>
#include <iostream>
#include <oaidl.h>
#include "MSOUTL.h"
#include <QDebug>

using TStringListPair = std::pair< QStringList, QStringList >;

COutlookAPI::EAddressTypes getAddressTypes( bool smtpOnly )
{
    return smtpOnly ? COutlookAPI::EAddressTypes::eSMTPOnly : COutlookAPI::EAddressTypes::eNone;
}

COutlookAPI::EAddressTypes getAddressTypes( std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly )
{
    auto types = getAddressTypes( smtpOnly );
    if ( recipientType.has_value() )
    {
        if ( recipientType == Outlook::OlMailRecipientType::olOriginator )
            types = types | COutlookAPI::EAddressTypes::eOriginator;
        if ( recipientType == Outlook::OlMailRecipientType::olTo )
            types = types | COutlookAPI::EAddressTypes::eTo;
        if ( recipientType == Outlook::OlMailRecipientType::olCC )
            types = types | COutlookAPI::EAddressTypes::eCC;
        if ( recipientType == Outlook::OlMailRecipientType::olBCC )
            types = types | COutlookAPI::EAddressTypes::eBCC;
    }
    else
        types = types | COutlookAPI::EAddressTypes::eAllRecipients;

    return types;
}

TStringListPair cleanResults( TStringListPair values )
{
    values.first.removeAll( QString() );
    values.second.removeAll( QString() );
    values.first.removeDuplicates();
    values.second.removeDuplicates();
    return values;
}

TStringListPair mergeResults( TStringListPair lhs, const TStringListPair &rhs )
{
    lhs.first << rhs.first;
    lhs.second << rhs.second;
    return cleanResults( lhs );
}

TStringListPair COutlookAPI::getEmailAddresses( Outlook::MailItem *mailItem, EAddressTypes types )   // returns the list of email addresses, display names
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

TStringListPair COutlookAPI::getEmailAddresses( std::shared_ptr< Outlook::MailItem > &mailItem, EAddressTypes types )   // returns the list of email addresses, display names
{
    if ( !mailItem )
        return {};
    return getEmailAddresses( mailItem.get(), types );
}

TStringListPair COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, EAddressTypes types )
{
    if ( !address )
        return {};
    auto type = address->AddressEntryUserType();

    bool smtpOnly = ( types & COutlookAPI::EAddressTypes::eSMTPOnly ) != 0;

    TStringListPair retVal;

    if ( address->GetExchangeUser() && !smtpOnly )
    {
        retVal.first << address->GetExchangeUser()->PrimarySmtpAddress();
        retVal.second << address->GetExchangeUser()->Name();
    }
    else if ( !smtpOnly && address->GetExchangeDistributionList() )
    {
        retVal.first << address->GetExchangeDistributionList()->PrimarySmtpAddress();
        retVal.second << address->GetExchangeDistributionList()->Name();
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

TStringListPair COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, EAddressTypes types )
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

TStringListPair COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, EAddressTypes types )
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

TStringListPair COutlookAPI::getEmailAddresses( Outlook::Recipients *recipients, EAddressTypes types )
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

TStringListPair COutlookAPI::getEmailAddresses( Outlook::Recipient *recipient, EAddressTypes types )
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

QStringList COutlookAPI::getSenderEmailAddresses( Outlook::MailItem *mailItem )
{
    return getEmailAddresses( mailItem, EAddressTypes::eSender ).first;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, bool smtpOnly /*= false*/ )
{
    return getEmailAddresses( address, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ).first;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, bool smtpOnly )
{
    return getEmailAddresses( entries, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ).first;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, bool smtpOnly )
{
    return getEmailAddresses( addresses, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ).first;
}

QString COutlookAPI::getEmailAddress( Outlook::Recipient *recipient, bool smtpOnly )
{
    auto retVal = getEmailAddresses( recipient, EAddressTypes::eAllEmailAddresses | getAddressTypes( smtpOnly ) ).first;
    if ( !retVal.empty() )
        return retVal.front();
    return {};
}

QStringList COutlookAPI::getRecipientEmails( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType, bool smtpOnly )
{
    return getEmailAddresses( mailItem->Recipients(),  getAddressTypes( recipientType, smtpOnly ) ).first;
}

QStringList COutlookAPI::getRecipientEmails( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly )
{
    return getEmailAddresses( recipients, getAddressTypes( recipientType, smtpOnly ) ).first;
}