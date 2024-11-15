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

    connect( fImpl->actionSelectServer, &QAction::triggered, this, &CMainWindow::slotSelectServer );

    connect( fImpl->actionReloadAllData, &QAction::triggered, this, &CMainWindow::slotReloadAll );
    connect( fImpl->actionReloadEmail, &QAction::triggered, this, &CMainWindow::slotReloadEmail );
    connect( fImpl->actionReloadFolders, &QAction::triggered, this, &CMainWindow::slotReloadFolders );
    connect( fImpl->actionReloadRules, &QAction::triggered, this, &CMainWindow::slotReloadRules );

    connect( fImpl->actionSortRules, &QAction::triggered, this, &CMainWindow::slotSortRules );
    connect( fImpl->actionRenameRules, &QAction::triggered, this, &CMainWindow::slotRenameRules );

    connect( fImpl->actionSelectServerAndRootFolder, &QAction::triggered, this, &CMainWindow::slotSelectServerAndInbox );
    connect( fImpl->actionAddRule, &QAction::triggered, this, &CMainWindow::slotAddRule );
    connect( fImpl->actionRunRule, &QAction::triggered, this, &CMainWindow::slotRunRule );
    connect( fImpl->actionAddToCurrentRule, &QAction::triggered, this, &CMainWindow::slotAddToCurrentRule );

    connect( fImpl->actionOnlyGroupUnread, &QAction::changed, [ = ]() { fImpl->email->setOnlyGroupUnread( fImpl->actionOnlyGroupUnread->isChecked() ); } );

    connect( fImpl->folders, &CFoldersView::sigFolderSelected, this, &CMainWindow::slotUpdateActions );
    connect( fImpl->email, &CEmailView::sigRuleSelected, this, &CMainWindow::slotUpdateActions );
    connect( fImpl->rules, &CRulesView::sigRuleSelected, this, &CMainWindow::slotUpdateActions );

    connect( fImpl->folders, &CFoldersView::sigFinishedLoading, [ = ]() { fImpl->rules->reload( true ); } );
    connect( fImpl->rules, &CRulesView::sigFinishedLoading, [ = ]() { fImpl->email->reload( true ); } );

    setWindowTitle( QObject::tr( "Rules Maker" ) );

    connect(
        COutlookHelpers::getInstance().get(), &COutlookHelpers::sigAccountChanged,
        [ = ]()
        {
            slotUpdateActions();
            slotReloadAll();
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
    bool accountSelected = COutlookHelpers::getInstance()->accountSelected();
    fImpl->actionReloadAllData->setEnabled( accountSelected );
    fImpl->actionReloadEmail->setEnabled( accountSelected );
    fImpl->actionReloadFolders->setEnabled( accountSelected );
    fImpl->actionReloadRules->setEnabled( accountSelected );
    fImpl->actionSortRules->setEnabled( accountSelected );
    fImpl->actionRenameRules->setEnabled( accountSelected );

    bool emailSelected = !fImpl->email->getRulesForSelection().isEmpty();
    bool ruleSelected = fImpl->rules->ruleSelected();
    bool folderSelected = !fImpl->folders->selectedPath().isEmpty();

    fImpl->actionRunRule->setEnabled( ruleSelected );
    fImpl->actionAddToCurrentRule->setEnabled( emailSelected && ruleSelected );
    fImpl->actionAddRule->setEnabled( accountSelected && folderSelected && emailSelected );
}

void CMainWindow::slotAddRule()
{
    auto destFolder = fImpl->folders->selectedFullPath();
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !COutlookHelpers::getInstance()->addRule( destFolder, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    fImpl->rules->reload( false );
}

void CMainWindow::slotAddToCurrentRule()
{
    auto rule = fImpl->rules->currentRule();
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !COutlookHelpers::getInstance()->addToRule( rule, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    fImpl->rules->reload( false );
}

void CMainWindow::slotRenameRules()
{
    COutlookHelpers::getInstance()->renameRules();
    slotReloadRules();
}

void CMainWindow::slotSortRules()
{
    COutlookHelpers::getInstance()->sortRules();
    slotReloadRules();
}

void CMainWindow::slotRunRule()
{
    fImpl->rules->runSelectedRule();
    fImpl->email->reload( false );
}

void CMainWindow::slotReloadAll()
{
    clearViews();
    if ( COutlookHelpers::getInstance()->accountSelected() )
        fImpl->folders->reload( true );
    slotUpdateActions();
}

void CMainWindow::slotReloadEmail()
{
    fImpl->email->clear();
    if ( COutlookHelpers::getInstance()->accountSelected() )
        fImpl->email->reload( false );
    slotUpdateActions();
}

void CMainWindow::slotReloadFolders()
{
    fImpl->folders->clear();
    if ( COutlookHelpers::getInstance()->accountSelected() )
        fImpl->folders->reload( false );
    slotUpdateActions();
}

void CMainWindow::slotReloadRules()
{
    fImpl->rules->clear();
    if ( COutlookHelpers::getInstance()->accountSelected() )
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
    auto account = COutlookHelpers::getInstance()->selectAccount( false, this );
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
