#include "OutlookAPI.h"

#include "SelectAccount.h"
#include <QSettings>
#include <QDebug>

#include "MSOUTL.h"

QString COutlookAPI::defaultProfileName() const
{
    if ( !fOutlookApp )
        return {};
    if ( fOutlookApp->isNull() )
        return {};
    return fOutlookApp->DefaultProfileName();
}

QString COutlookAPI::accountName() const
{
    if ( !accountSelected() )
        return {};
    return fAccount->DisplayName();
}

bool COutlookAPI::accountSelected() const
{
    return fAccount.operator bool() && !fAccount->isNull();
}

COutlookObj< Outlook::_Account > COutlookAPI::closeAndSelectAccount( bool notifyOnChange )
{
    logout( notifyOnChange );
    return selectAccount( notifyOnChange );
}

COutlookObj< Outlook::_NameSpace > COutlookAPI::getNamespace( Outlook::_NameSpace *ns )
{
    if ( !ns )
        return {};
    return COutlookObj< Outlook::_NameSpace >( ns );
}

std::optional< std::map< QString, COutlookObj< Outlook::_Account > > > COutlookAPI::getAllAccounts( const QString &profileName )
{
    if ( !connected() )
    {
        fSession = getNamespace( fOutlookApp->Session() );
        if ( !fSession )
            return {};

        if ( !profileName.isEmpty() )
            fSession->Logon( profileName );
        else
            fSession->Logon();
        fLoggedIn = true;
    }

    auto accounts = fSession->Accounts();
    if ( !accounts )
    {
        return {};
    }

    std::map< QString, COutlookObj< Outlook::_Account > > retVal;
    auto numAccounts = accounts->Count();

    for ( auto ii = 1; ii <= numAccounts; ++ii )
    {
        auto item = accounts->Item( ii );
        if ( !item )
            continue;

        auto account = getAccount( item );

        if ( account->AccountType() != Outlook::OlAccountType::olExchange )
            continue;

        switch ( account->ExchangeConnectionMode() )
        {
            case Outlook::OlExchangeConnectionMode::olNoExchange:
            case Outlook::OlExchangeConnectionMode::olOffline:
            case Outlook::OlExchangeConnectionMode::olCachedOffline:
            case Outlook::OlExchangeConnectionMode::olDisconnected:
            case Outlook::OlExchangeConnectionMode::olCachedDisconnected:
                continue;
            default:
                break;
        }

        retVal[ account->DisplayName() ] = account;
    }
    return retVal;
}

QString COutlookAPI::defaultAccountName()
{
    auto lastAccount = lastAccountName();
    auto allAccounts = getAllAccounts( lastAccount );
    if ( !allAccounts.has_value() )
        return {};

    if ( allAccounts.value().size() == 1 )
    {
        return ( *allAccounts.value().begin() ).first;
    }
    if ( !lastAccount.isEmpty() )
    {
        if ( allAccounts.value().find( lastAccount ) != allAccounts.value().end() )
            return lastAccount;
    }

    return {};
}

bool COutlookAPI::connected()
{
    if ( !fLoggedIn )
        return false;
    if ( !fSession || fSession->isNull() )
        return false;
    return ( fSession->ExchangeConnectionMode() != Outlook::OlExchangeConnectionMode::olOffline );
}

bool COutlookAPI::selectAccount( const QString &accountName, bool notifyOnChange )
{
    if ( !connected() )
        return false;

    auto accounts = getAllAccounts( {} );
    auto pos = accounts.value().find( accountName );
    if ( pos == accounts.value().end() )
        return false;

    QSettings settings;
    settings.setValue( "Account", accountName );
    fAccount = ( *pos ).second;

    if ( notifyOnChange )
        emit sigAccountChanged();

    return accountSelected();
}

COutlookObj< Outlook::_Account > COutlookAPI::selectAccount( bool notifyOnChange )
{
    if ( accountSelected() )
        return fAccount;

    if ( fOutlookApp->isNull() )
        return {};

    CSelectAccount dlg( getParentWidget() );
    auto init = dlg.initResult();
    if ( ( init == CSelectAccount::EInitResult::eError ) || ( init == CSelectAccount::EInitResult::eNoAccounts ) )
    {
        logout( notifyOnChange );
        return {};
    }

    if ( init != CSelectAccount::EInitResult::eSingleAccount )
    {
        if ( dlg.exec() != QDialog::Accepted )
        {
            return {};
        }
    }

    auto &&[ accountName, account ] = dlg.account();
    fAccount = dlg.account().second;
    if ( !fAccount )
    {
        logout( notifyOnChange );
        return {};
    }
    setLastAccountName( fAccount->DisplayName() );
    setLoadAccountInfo( dlg.loadAccountInfo() );

    if ( notifyOnChange )
        emit sigAccountChanged();

    if ( !accountSelected() )
        return {};
    return fAccount;
}

COutlookObj< Outlook::_Account > COutlookAPI::getAccount( Outlook::_Account *item )
{
    if ( !item )
        return {};

    return COutlookObj< Outlook::_Account >( item );
}
