#include "OutlookAPI.h"

#include "SelectAccount.h"
#include <QSettings>

#include "OutlookLib/MSOUTL.h"
#include <objbase.h>

#include <iostream>

bool COutlookAPI::connected()
{
    if ( !fLoggedIn || !fOutlookApp )
        return false;
    if ( !fSession || fSession->isNull() )
        return false;
    return ( static_cast< Outlook::OlExchangeConnectionMode >( fSession->ExchangeConnectionMode() ) != Outlook::OlExchangeConnectionMode::olOffline );
}

bool COutlookAPI::logon( const QString &profileName )
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
    return fLoggedIn;
}

void COutlookAPI::logout( bool andNotify )
{
    fSession.reset();
    fAccount.reset();
    fInbox.reset();
    fRootFolder.reset();
    fJunkFolder.reset();
    fTrashFolder.reset();
    fContacts.reset();
    fRules.reset();

    if ( fLoggedIn && fOutlookApp && !fOutlookApp->isNull() && fOutlookApp->Session() )
    {
        Outlook::NameSpace( fOutlookApp->Session() ).Logoff();
        if ( !outlookProcessRunning() )
            fOutlookApp->Quit();
        fLoggedIn = false;
        if ( andNotify )
            emit sigAccountChanged();
    }
    CoUninitialize();
}

bool COutlookAPI::outlookFullySetup( bool resetOutlook )
{
    if ( !fOutlookApp )
        return false;
    if ( fOutlookApp->isNull() )
        return false;
    auto profileName = defaultProfileName();
    auto accountName = defaultAccountName();
    //std::cout << "Profile Name: '" << qPrintable( profileName  ) << "'\n";
    //std::cout << "Account Name: '" << qPrintable( accountName ) << "'" << std::endl;
    if ( resetOutlook )
        this->resetApplication();
    return !profileName.isEmpty() && !accountName.isEmpty();
}

QString COutlookAPI::defaultProfileName() const
{
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

std::shared_ptr< Outlook::Account > COutlookAPI::closeAndSelectAccount( bool notifyOnChange )
{
    logout( notifyOnChange );
    return selectAccount( notifyOnChange );
}

std::shared_ptr< Outlook::NameSpace > COutlookAPI::getNamespace( Outlook::_NameSpace *ns )
{
    if ( !ns )
        return {};
    return connectToException( std::make_shared< Outlook::NameSpace >( ns ) );
}

std::optional< std::map< QString, std::shared_ptr< Outlook::Account > > > COutlookAPI::getAllAccounts( const QString &profileName )
{
    if ( !logon( profileName ) )
        return {};

    auto accounts = fSession->Accounts();
    if ( !accounts )
    {
        return {};
    }

    std::map< QString, std::shared_ptr< Outlook::Account > > retVal;
    auto numAccounts = accounts->Count();

    for ( auto ii = 1; ii <= numAccounts; ++ii )
    {
        auto item = accounts->Item( ii );
        if ( !item )
            continue;

        auto account = getAccount( item );

        if ( static_cast< Outlook::OlAccountType >( account->AccountType() ) != Outlook::OlAccountType::olExchange )
            continue;

        switch ( static_cast< Outlook::OlExchangeConnectionMode >( account->ExchangeConnectionMode() ) )
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
    auto profileName = defaultProfileName();
    if ( profileName.isEmpty() )
        return {};

    auto allAccounts = getAllAccounts( profileName );
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

std::shared_ptr< Outlook::Account > COutlookAPI::selectAccount( bool notifyOnChange )
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

std::shared_ptr< Outlook::Account > COutlookAPI::getAccount( Outlook::_Account *item )
{
    if ( !item )
        return {};

    return connectToException( std::make_shared< Outlook::Account >( item ) );
}
