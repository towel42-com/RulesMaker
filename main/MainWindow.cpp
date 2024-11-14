#include "MainWindow.h"
#include "OutlookHelpers.h"
#include "OutlookSetup.h"

#include "ui_MainWindow.h"

#include <QTimer>
#include <QMessageBox>

CMainWindow::CMainWindow( QWidget *parent ) :
    QMainWindow( parent ),
    fImpl( new Ui::CMainWindow )
{
    fImpl->setupUi( this );

    connect( fImpl->actionSelectServer, &QAction::triggered, [ = ]() { slotSelectServer(); } );
    connect( fImpl->actionReloadData, &QAction::triggered, [ = ]() { slotReload(); } );
    connect( fImpl->actionSelectServerAndInbox, &QAction::triggered, [ = ]() { slotSelectServerAndInbox(); } );
    connect( fImpl->actionAddRule, &QAction::triggered, [ = ]() { slotAddRule(); } );
    connect( fImpl->actionOnlyGroupUnread, &QAction::changed, [ = ]() { fImpl->email->setOnlyGroupUnread( fImpl->actionOnlyGroupUnread->isChecked() ); } );

    connect( fImpl->folders, &CFoldersView::sigFinishedLoading, [ = ]() { fImpl->rules->reload(); } );
    connect( fImpl->rules, &CRulesView::sigFinishedLoading, [ = ]() { fImpl->email->reload(); } );

    connect( fImpl->folders, &CFoldersView::sigFolderSelected, [ = ]() { slotUpdateActions(); } );
    connect( fImpl->email, &CEmailView::sigRuleSelected, [ = ]() { slotUpdateActions(); } );

    setWindowTitle( QObject::tr( "Rules Maker" ) );

    connect(
        COutlookHelpers::getInstance().get(), &COutlookHelpers::sigAccountChanged,
        [ = ]()
        {
            slotUpdateActions();
            slotReload();
        } );
    slotUpdateActions();
    fImpl->actionOnlyGroupUnread->setChecked( fImpl->email->onlyGroupUnread() );
}

CMainWindow::~CMainWindow()
{
    clearViews();
    COutlookHelpers::getInstance()->logout( false );
}

void CMainWindow::slotUpdateActions()
{
    fImpl->actionReloadData->setEnabled( COutlookHelpers::getInstance()->accountSelected() );
    bool allowAddRule = COutlookHelpers::getInstance()->accountSelected();
    allowAddRule &= !fImpl->folders->currentPath().isEmpty();
    allowAddRule &= !fImpl->email->currentRule().isEmpty();
    fImpl->actionAddRule->setEnabled( allowAddRule );
}

void CMainWindow::slotAddRule()
{
    auto destFolder = fImpl->folders->fullPath();
    auto rules = fImpl->email->currentRule();

    QStringList msgs;
    if ( !COutlookHelpers::getInstance()->addRule( destFolder, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
}

void CMainWindow::slotReload()
{
    clearViews();
    if ( COutlookHelpers::getInstance()->accountSelected() )
        fImpl->folders->reload();
    slotUpdateActions();
}

void CMainWindow::clearViews()
{
    fImpl->email->clear();
    fImpl->rules->clear();
    fImpl->folders->clear();
}

void CMainWindow::slotSelectServer()
{
    clearViews();
    auto account = COutlookHelpers::getInstance()->selectAccount( false, this );
    if ( !account )
        return;
    slotReload();
}

void CMainWindow::slotSelectServerAndInbox()
{
    clearViews();
    COutlookSetup dlg;
    if ( dlg.exec() == QDialog::Accepted )
        slotReload();
    slotReload();
}
