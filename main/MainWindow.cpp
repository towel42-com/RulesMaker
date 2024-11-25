#include "MainWindow.h"
#include "OutlookAPI.h"
#include "OutlookSetup.h"

#include "ui_MainWindow.h"

#include <QSettings>
#include <QTimer>
#include <QMessageBox>

CMainWindow::CMainWindow( QWidget *parent ) :
    QMainWindow( parent ),
    fImpl( new Ui::CMainWindow )
{
    fImpl->setupUi( this );

    connect( fImpl->actionSelectServer, &QAction::triggered, this, &CMainWindow::slotSelectServer );

    connect( fImpl->actionReloadAllData, &QAction::triggered, this, &CMainWindow::slotReloadAll );
    connect( fImpl->actionReloadEmail, &QAction::triggered, this, &CMainWindow::slotReloadEmail );
    connect( fImpl->actionReloadFolders, &QAction::triggered, this, &CMainWindow::slotReloadFolders );
    connect( fImpl->actionReloadRules, &QAction::triggered, this, &CMainWindow::slotReloadRules );

    connect( fImpl->actionSortRules, &QAction::triggered, this, &CMainWindow::slotSortRules );
    connect( fImpl->actionRenameRules, &QAction::triggered, this, &CMainWindow::slotRenameRules );
    connect( fImpl->actionMergeRules, &QAction::triggered, this, &CMainWindow::slotMergeRules );
    connect( fImpl->actionMoveFromToAddress, &QAction::triggered, this, &CMainWindow::slotMoveFromToAddress );

    connect( fImpl->actionSelectServerAndRootFolder, &QAction::triggered, this, &CMainWindow::slotSelectServerAndInbox );
    connect( fImpl->actionAddRule, &QAction::triggered, this, &CMainWindow::slotAddRule );
    connect( fImpl->actionRunRule, &QAction::triggered, this, &CMainWindow::slotRunRule );
    connect( fImpl->actionAddToSelectedRule, &QAction::triggered, this, &CMainWindow::slotAddToSelectedRule );

    connect(
        fImpl->actionProcessAllEmailWhenLessThan200Emails, &QAction::changed,
        [ = ]()
        {
            QSettings settings;
            settings.setValue( "ProcessAllEmailWhenLessThan200Emails", fImpl->actionProcessAllEmailWhenLessThan200Emails->isChecked() );
            fImpl->email->setProcessAllEmailWhenLessThan200Emails( fImpl->actionProcessAllEmailWhenLessThan200Emails->isChecked() );
        } );

    connect(
        fImpl->actionOnlyProcessUnread, &QAction::changed,
        [ = ]()
        {
            QSettings settings;
            settings.setValue( "OnlyProcessUnread", fImpl->actionOnlyProcessUnread->isChecked() );
            fImpl->email->setOnlyProcessUnread( fImpl->actionOnlyProcessUnread->isChecked() );
        } );

    connect( fImpl->folders, &CFoldersView::sigFolderSelected, this, &CMainWindow::slotUpdateActions );
    connect( fImpl->email, &CEmailView::sigRuleSelected, this, &CMainWindow::slotUpdateActions );
    connect( fImpl->rules, &CRulesView::sigRuleSelected, this, &CMainWindow::slotUpdateActions );

    connect( fImpl->folders, &CFoldersView::sigFinishedLoading, [ = ]() { fImpl->rules->reload( true ); } );
    connect( fImpl->rules, &CRulesView::sigFinishedLoading, [ = ]() { fImpl->email->reload( true ); } );

    setWindowTitle( QObject::tr( "Rules Maker" ) );

    connect(
        COutlookAPI::getInstance().get(), &COutlookAPI::sigAccountChanged,
        [ = ]()
        {
            slotUpdateActions();
            slotReloadAll();
        } );
    slotUpdateActions();

    QSettings settings;
    fImpl->email->setOnlyProcessUnread( settings.value( "OnlyProcessUnread", true ).toBool() );
    fImpl->email->setProcessAllEmailWhenLessThan200Emails( settings.value( "ProcessAllEmailWhenLessThan200Emails", true ).toBool() );
    settings.setValue( "ProcessAllEmailWhenLessThan200Emails", fImpl->actionProcessAllEmailWhenLessThan200Emails->isChecked() );
    fImpl->actionProcessAllEmailWhenLessThan200Emails->setChecked( fImpl->email->processAllEmailWhenLessThan200Emails() );
    fImpl->actionOnlyProcessUnread->setChecked( fImpl->email->onlyProcessUnread() );

    QTimer::singleShot( 0, [ = ]() { slotSelectServer(); } );
}

CMainWindow::~CMainWindow()
{
    clearViews();
    COutlookAPI::getInstance()->logout( false );
}

void CMainWindow::slotUpdateActions()
{
    bool accountSelected = COutlookAPI::getInstance()->accountSelected();
    fImpl->actionReloadAllData->setEnabled( accountSelected );
    fImpl->actionReloadEmail->setEnabled( accountSelected );
    fImpl->actionReloadFolders->setEnabled( accountSelected );
    fImpl->actionReloadRules->setEnabled( accountSelected );
    fImpl->actionSortRules->setEnabled( accountSelected );
    fImpl->actionRenameRules->setEnabled( accountSelected );
    fImpl->actionMoveFromToAddress->setEnabled( accountSelected );

    bool emailSelected = !fImpl->email->getRulesForSelection().isEmpty();
    bool ruleSelected = fImpl->rules->ruleSelected();
    bool folderSelected = !fImpl->folders->selectedPath().isEmpty();

    fImpl->actionRunRule->setEnabled( ruleSelected );
    fImpl->actionAddToSelectedRule->setEnabled( emailSelected && ruleSelected );
    fImpl->actionAddRule->setEnabled( accountSelected && folderSelected && emailSelected );
}

void CMainWindow::slotAddRule()
{
    auto destFolder = fImpl->folders->selectedFullPath();
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !fImpl->rules->addRule( destFolder, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    slotReloadEmail();
}

void CMainWindow::slotAddToSelectedRule()
{
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !fImpl->rules->addToSelectedRule( rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    slotReloadEmail();
}

void CMainWindow::slotMergeRules()
{
    COutlookAPI::getInstance()->mergeRules();
    slotReloadRules();
}

void CMainWindow::slotRenameRules()
{
    COutlookAPI::getInstance()->renameRules();
    slotReloadRules();
}

void CMainWindow::slotSortRules()
{
    COutlookAPI::getInstance()->sortRules();
    slotReloadRules();
}

void CMainWindow::slotMoveFromToAddress()
{
    COutlookAPI::getInstance()->moveFromToAddress();
    slotReloadRules();
}

void CMainWindow::slotRunRule()
{
    fImpl->rules->runSelectedRule();
    slotReloadEmail();
}

void CMainWindow::slotReloadAll()
{
    clearViews();
    if ( COutlookAPI::getInstance()->accountSelected() )
    {
        fImpl->folders->reload( true );
        setWindowTitle( tr( "Outlook Rules Maker - %1" ).arg( COutlookAPI::getInstance()->accountName() ) );
    }
    slotUpdateActions();
}

void CMainWindow::slotReloadEmail()
{
    fImpl->email->clear();
    if ( COutlookAPI::getInstance()->accountSelected() )
        fImpl->email->reload( false );
    slotUpdateActions();
}

void CMainWindow::slotReloadFolders()
{
    fImpl->folders->clear();
    if ( COutlookAPI::getInstance()->accountSelected() )
        fImpl->folders->reload( false );
    slotUpdateActions();
}

void CMainWindow::slotReloadRules()
{
    fImpl->rules->clear();
    if ( COutlookAPI::getInstance()->accountSelected() )
        fImpl->rules->reload( false );
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
    auto account = COutlookAPI::getInstance()->selectAccount( false, this );
    if ( !account )
        return;
    slotReloadAll();
}

void CMainWindow::slotSelectServerAndInbox()
{
    clearViews();
    COutlookSetup dlg;
    if ( dlg.exec() == QDialog::Accepted )
        slotReloadAll();
    slotReloadAll();
}
