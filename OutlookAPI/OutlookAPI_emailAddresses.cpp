#include "OutlookAPI.h"
#include "EmailAddress.h"
#include "OutlookObj.h"

#include "MSOUTL.h"

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

TEmailAddressList COutlookAPI::getEmailAddresses( const COutlookObj< Outlook::MailItem > &mailItem, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )   // returns the list of email addresses, display names
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
        retVal = getEmailAddresses( mailItem->Sender(), contactTypes );
        bool isExchangeContact = isExchangeUser( mailItem->Sender() );
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

QStringList COutlookAPI::getSenderEmailAddresses( Outlook::MailItem *mailItem )
{
    return getAddresses( getEmailAddresses( mailItem, EAddressTypes::eSender ) );
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
        auto curr = getEmailAddresses( recipient, types, contactTypes );
        retVal.insert( retVal.end(), curr.begin(), curr.end() );
    }
    return retVal;
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::Recipient *recipient, std::optional< EAddressTypes > types, std::optional< EContactTypes > contactTypes )
{
    if ( !recipient )
        return {};

    if ( !isAddressType( static_cast< Outlook::OlMailRecipientType >( recipient->Type() ), types ) )
        return {};

    auto retVal = getEmailAddresses( recipient->AddressEntry(), contactTypes );
    if ( retVal.empty() )
    {
        bool isExchangeContact = isExchangeUser( recipient->AddressEntry() );
        if ( isContactType( isExchangeContact, contactTypes ) )
            retVal.push_back( std::make_shared< CEmailAddress >( recipient->Address(), recipient->Name(), isExchangeContact ) );
    }

    return retVal;
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses, std::optional< EContactTypes > contactTypes )
{
    if ( !addresses )
        return {};

    return getEmailAddresses( addresses->AddressEntries(), contactTypes );
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries, std::optional< EContactTypes > contactTypes )
{
    if ( !entries )
        return {};

    TEmailAddressList retVal;
    auto num = entries->Count();
    for ( int ii = 1; ii <= num; ++ii )
    {
        auto entry = entries->Item( ii );
        if ( !entry )
            continue;
        auto curr = getEmailAddresses( entry, contactTypes );
        retVal.insert( retVal.end(), curr.begin(), curr.end() );
    }
    return cleanResults( retVal );
}

TEmailAddressList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address, std::optional< EContactTypes > contactTypes )
{
    if ( !address )
        return {};
    if ( !isContactType( address->AddressEntryUserType(), contactTypes ) )
        return {};

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
