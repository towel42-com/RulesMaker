#include "Settings.h"
#include "OutlookAPI/OutlookAPI.h"

#include "ui_Settings.h"

#include <QInputDialog>

CSettings::CSettings( QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::CSettings )
{
    init();
}

void CSettings::init()
{
    fImpl->setupUi( this );

    auto api = COutlookAPI::instance();
    fImpl->onlyProcessUnread->setChecked( api->onlyProcessUnread() );
    fImpl->processAllEmailWhenLessThan200Emails->setChecked( api->processAllEmailWhenLessThan200Emails() );
    fImpl->onlyProcessTheFirst500Emails->setChecked( api->onlyProcessTheFirst500Emails() );
    fImpl->disableRatherThanDeleteRules->setChecked( api->disableRatherThanDeleteRules() );
    fImpl->rulesToSkip->addItems( api->rulesToSkip() );
    fImpl->loadAccountInfo->setChecked( api->loadAccountInfo() );

    connect( fImpl->rulesToSkip, &QListWidget::itemSelectionChanged, this, &CSettings::slotRegexSelectionChanged );
    connect( fImpl->addRegex, &QAbstractButton::clicked, this, &CSettings::slotAddRegex );
    connect( fImpl->delRegex, &QAbstractButton::clicked, this, &CSettings::slotDelRegex );

    slotRegexSelectionChanged();
}

CSettings::~CSettings()
{
}

void CSettings::accept()
{
    auto api = COutlookAPI::instance();
    api->setOnlyProcessUnread( fImpl->onlyProcessUnread->isChecked() );
    api->setProcessAllEmailWhenLessThan200Emails( fImpl->processAllEmailWhenLessThan200Emails->isChecked() );
    api->setOnlyProcessTheFirst500Emails( fImpl->onlyProcessTheFirst500Emails->isChecked() );
    api->setLoadAccountInfo( fImpl->loadAccountInfo->isChecked() );

    api->setDisableRatherThanDeleteRules( fImpl->disableRatherThanDeleteRules->isChecked() );
    auto regexes = QStringList();
    for ( auto &&ii = 0; ii < fImpl->rulesToSkip->count(); ++ii )
    {
        regexes << fImpl->rulesToSkip->item( ii )->text();
    }
    api->setRulesToSkip( regexes );
    QDialog::accept();
}

bool CSettings::changed() const
{
    auto api = COutlookAPI::instance();
    bool retVal = fImpl->onlyProcessUnread->isChecked() != api->onlyProcessUnread();
    retVal = retVal || ( fImpl->processAllEmailWhenLessThan200Emails->isChecked() != api->processAllEmailWhenLessThan200Emails() );
    retVal = retVal || ( fImpl->onlyProcessTheFirst500Emails->isChecked() != api->onlyProcessTheFirst500Emails() );
    retVal = retVal || ( fImpl->loadAccountInfo->isChecked() != api->loadAccountInfo() );

    retVal = retVal || ( fImpl->disableRatherThanDeleteRules->isChecked() != api->disableRatherThanDeleteRules() );

    auto prevRegexes = api->rulesToSkip();
    retVal = retVal || ( prevRegexes.size() != fImpl->rulesToSkip->count() );
    for ( auto ii = 0; !retVal && ii < fImpl->rulesToSkip->count(); ++ii )
    {
        retVal = retVal || ( fImpl->rulesToSkip->item( ii )->text() != prevRegexes[ ii ] );
    }
    return retVal;
}

void CSettings::slotRegexSelectionChanged()
{
    fImpl->delRegex->setEnabled( fImpl->rulesToSkip->selectedItems().size() > 0 );
}

void CSettings::slotAddRegex()
{
    auto regex = QInputDialog::getText( this, tr( "Add Regex" ), tr( "Enter a regex to skip:" ) );
    if ( !regex.isEmpty() )
    {
        fImpl->rulesToSkip->addItem( regex );
    }
}

void CSettings::slotDelRegex()
{
    auto items = fImpl->rulesToSkip->selectedItems();
    if ( items.empty() )
        return;
    for ( auto &&ii : items )
    {
        delete ii;
    }
}
