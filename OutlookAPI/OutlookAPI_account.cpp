#include "OutlookAPI.h"

#include <QInputDialog>
#include <QSettings>
#include "MSOUTL.h"

QString COutlookAPI::accountName() const
{
    if ( !accountSelected() )
        return {};
    return fAccount->DisplayName();
}

bool COutlookAPI::accountSelected() const
{
    return fAccount.operator bool();
}

std::shared_ptr< Outlook::Account > COutlookAPI::selectAccount( bool notifyOnChange )
{
    logout( notifyOnChange );
    if ( fOutlookApp->isNull() )
        return {};

    auto profileName = fOutlookApp->DefaultProfileName();
    Outlook::NameSpace session( fOutlookApp->Session() );
    session.Logon( profileName );
    fLoggedIn = true;

    std::vector< std::shared_ptr< Outlook::Account > > allAccounts;

    auto accounts = session.Accounts();
    if ( !accounts )
    {
        logout( notifyOnChange );
        return {};
    }

    auto numAccounts = accounts->Count();
    allAccounts.reserve( numAccounts );

    QSettings settings;
    auto lastAccount = settings.value( "Account", QString() ).toString();
    int accountPos = 0;
    QStringList accountNames;
    std::map< QString, std::shared_ptr< Outlook::Account > > accountMap;
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

        allAccounts.push_back( account );
        auto accountName = account->DisplayName();
        accountNames << accountName;
        accountMap[ accountName ] = account;
        if ( accountName == lastAccount )
        {
            accountPos = ii - 1;
        }
    }

    if ( allAccounts.size() == 0 )
    {
        logout( notifyOnChange );
        return {};
    }

    if ( allAccounts.size() == 1 )
    {
        fAccount = allAccounts.front();
        settings.setValue( "Account", allAccounts.front()->DisplayName() );
        return fAccount;
    }

    bool aOK{ false };
    auto account = QInputDialog::getItem( fParentWidget, QString( "Select Account:" ), "Account:", accountNames, accountPos, false, &aOK );
    if ( !aOK || account.isEmpty() )
    {
        logout( notifyOnChange );
        return {};
    }
    auto pos = accountMap.find( account );
    if ( pos == accountMap.end() )
    {
        logout( notifyOnChange );
        return {};
    }
    settings.setValue( "Account", account );
    fAccount = ( *pos ).second;

    if ( notifyOnChange )
        emit sigAccountChanged();
    return fAccount;
}


std::shared_ptr< Outlook::Account > COutlookAPI::getAccount( Outlook::_Account *item )
{
    if ( !item )
        return {};

    return connectToException( std::make_shared< Outlook::Account >( item ) );
}
