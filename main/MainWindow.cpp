#include "MainWindow.h"
#include "OutlookAPI.h"
#include "OutlookSetup.h"
#include "StatusProgress.h"
#include "MSOUTL.h"

#include "ui_MainWindow.h"

#include <QSettings>
#include <QTimer>
#include <QMessageBox>
#include <QPushButton>
#include <QCursor>
#include <QApplication>

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

    setupStatusBar();

    setWindowTitle( QObject::tr( "Rules Maker" ) );

    auto api = COutlookAPI::getInstance();
    connect(
        api.get(), &COutlookAPI::sigAccountChanged,
        [ = ]()
        {
            slotUpdateActions();
            slotReloadAll();
        } );

    connect(
        api.get(), &COutlookAPI::sigInitStatus,
        [ = ]( const QString &label, int max )
        {
            CStatusProgress *bar = nullptr;
            auto pos = fProgressBars.find( label );
            if ( pos == fProgressBars.end() )
                bar = addStatusBar( label, nullptr, true );
            else
                bar = ( *pos ).second;
            bar->setRange( 0, max );
        } );
    connect(
        api.get(), &COutlookAPI::sigSetStatus,
        [ = ]( const QString &label, int curr, int max )
        {
            auto pos = fProgressBars.find( label );
            if ( pos == fProgressBars.end() )
                return;

            auto bar = ( *pos ).second;
            bar->slotSetStatus( curr, max );
        } );
    connect(
        api.get(), &COutlookAPI::sigIncStatusValue,
        [ = ]( const QString &label )
        {
            auto pos = fProgressBars.find( label );
            if ( pos == fProgressBars.end() )
                return;

            auto bar = ( *pos ).second;
            bar->slotIncValue();
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
    bool folderSame = false;
    if ( emailSelected && ruleSelected )
    {
        auto selectedFolder = fImpl->folders->selectedFolder();
        if ( selectedFolder )
        {
            auto ruleFolder = fImpl->rules->folderForSelectedRule();
            auto selectedFolderPath = selectedFolder->FullFolderPath();
            folderSame = ruleFolder == selectedFolderPath;
        }
        else
            folderSame = true;
    }

    fImpl->actionAddToSelectedRule->setEnabled( emailSelected && ruleSelected && folderSame );
    fImpl->actionAddRule->setEnabled( accountSelected && folderSelected && emailSelected );
}

void CMainWindow::slotAddRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto destFolder = fImpl->folders->selectedFullPath();
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !fImpl->rules->addRule( destFolder, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotAddToSelectedRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !fImpl->rules->addToSelectedRule( rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotMergeRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::getInstance()->mergeRules();
    slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRenameRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::getInstance()->renameRules();
    slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotSortRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::getInstance()->sortRules();
    slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotMoveFromToAddress()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::getInstance()->moveFromToAddress();
    slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    fImpl->rules->runSelectedRule();
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotReloadAll()
{
    clearViews();
    auto windowTitle = tr( "Outlook Rules Maker" );
    if ( COutlookAPI::getInstance()->accountSelected() )
    {
        fImpl->folders->reload( true );
        fImpl->rules->reload( true );
        fImpl->email->reload( true );
        windowTitle += tr( " - %1" ).arg( COutlookAPI::getInstance()->accountName() );
    }
    setWindowTitle( windowTitle );
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
}

void CMainWindow::slotHandleProgressToggle()
{
    bool visible = false;
    for ( auto &&ii : fProgressBars )
    {
        if ( ii.second->isVisible() )
            visible = true;
        if ( visible )
            break;
    }

    fCancelButton->setVisible( visible );
}

CStatusProgress *CMainWindow::addStatusBar( const QString &label, QObject *object, bool hasInc )
{
    auto progress = new CStatusProgress( label );
    if ( object )
    {
        connect( object, SIGNAL( sigSetStatus( int, int ) ), progress, SLOT( slotSetStatus( int, int ) ) );
        if ( hasInc )
            connect( object, SIGNAL( sigIncStatusValue() ), progress, SLOT( slotIncValue() ) );
    }
    connect( progress, &CStatusProgress::sigShow, this, &CMainWindow::slotHandleProgressToggle );
    connect( progress, &CStatusProgress::sigFinished, this, &CMainWindow::slotHandleProgressToggle );
    auto num = statusBar()->insertPermanentWidget( static_cast< int >( fProgressBars.size() ), progress );
    fProgressBars[ label ] = progress;
    return progress;
}

void CMainWindow::setupStatusBar()
{
    if ( !fProgressBars.empty() )
        return;

    addStatusBar( "Loading Folders:", fImpl->folders, true );
    addStatusBar( "Grouping Emails:", fImpl->email, false );
    addStatusBar( "Loading Rules:", fImpl->rules, false );

    fCancelButton = new QPushButton( "&Cancel" );
    connect( fCancelButton, &QPushButton::clicked, COutlookAPI::getInstance().get(), &COutlookAPI::slotCanceled );
    connect(
        fCancelButton, &QPushButton::clicked, this,
        [ = ]()
        {
            for ( auto &&ii : fProgressBars )
                ii.second->hide();
            fCancelButton->hide();
        } );
    statusBar()->addPermanentWidget( fCancelButton );

    for ( auto &&ii : fProgressBars )
        ii.second->hide();
    slotHandleProgressToggle();
}
