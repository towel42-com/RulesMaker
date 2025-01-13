#include "SelectAccount.h"
#include "OutlookAPI.h"

#include "ui_SelectAccount.h"

#include <QSettings>

CSelectAccount::CSelectAccount( QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::CSelectAccount )
{
    init();
}

void CSelectAccount::init()
{
    fImpl->setupUi( this );

    auto api = COutlookAPI::instance();

    fImpl->loadServer->setChecked( api->loadAccountInfo() );
    fAllAccounts = api->getAllAccounts( api->defaultAccountName() );
    if ( !fAllAccounts.has_value() )
    {
        fInitResult = EInitResult::eError;
        return;
    }

    if ( fAllAccounts.value().size() == 0 )
    {
        fInitResult = EInitResult::eNoAccounts;
        return;
    }

    if ( fAllAccounts.value().size() == 1 )
    {
        auto pos = fAllAccounts.value().begin();
        fAccount = *pos;
        fInitResult = EInitResult::eSingleAccount;
        return;
    }

    QStringList accountNames;
    int accountPos = -1;
    int currPos = 0;
    for ( auto &&account : fAllAccounts.value() )
    {
        accountNames << account.first;
        if ( account.first == api->defaultAccountName() )
        {
            accountPos = static_cast< int >( currPos );
        }
        currPos++;
    }
    fImpl->profiles->clear();
    fImpl->profiles->addItems( accountNames );
    fImpl->profiles->setCurrentIndex( accountPos );
    fInitResult = EInitResult::eSuccess;
}

CSelectAccount::~CSelectAccount()
{
}

std::pair< QString, std::shared_ptr< Outlook::Account > > CSelectAccount::account() const
{
    return fAccount;
}

bool CSelectAccount::loadAccountInfo() const
{
    return fImpl->loadServer->isChecked();
}

void CSelectAccount::accept()
{
    auto currText = fImpl->profiles->currentText();
    auto pos = fAllAccounts.value().find( currText );
    if ( pos != fAllAccounts.value().end() )
    {
        fAccount = *pos;
        QDialog::accept();
    }
    else
        QDialog::reject();
}
