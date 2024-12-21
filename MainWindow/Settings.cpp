#include "Settings.h"
#include "OutlookAPI/OutlookAPI.h"

#include "ui_Settings.h"

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
    fImpl->processAllEmailWhenLessThan200Emails->setChecked( api->processAllEmailWhenLessThan200Emails() );
    fImpl->onlyProcessUnread->setChecked( api->onlyProcessUnread() );
    fImpl->includeJunkFolderWhenRunningOnAllFolders->setChecked( api->includeJunkInRunAllFolders() );
    fImpl->disableRatherThanDeleteRules->setChecked( api->disableRatherThanDeleteRules() );
}

CSettings::~CSettings()
{
}

void CSettings::accept()
{
    auto api = COutlookAPI::instance();
    api->setProcessAllEmailWhenLessThan200Emails( fImpl->processAllEmailWhenLessThan200Emails->isChecked() );
    api->setOnlyProcessUnread( fImpl->onlyProcessUnread->isChecked() );
    api->setIncludeJunkInRunAllFolders( fImpl->includeJunkFolderWhenRunningOnAllFolders->isChecked() );
    api->setDisableRatherThanDeleteRules( fImpl->disableRatherThanDeleteRules->isChecked() );
    QDialog::accept();
}
