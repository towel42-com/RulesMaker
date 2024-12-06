#include "MainWindow.h"
#include "OutlookAPI.h"
#include "StatusProgress.h"
#include "MSOUTL.h"

#include "ui_MainWindow.h"

#include <QTimer>
#include <QMessageBox>
#include <QPushButton>
#include <QCursor>
#include <QApplication>
#include <QLineEdit>
#include <QToolButton>

CMainWindow::CMainWindow( QWidget *parent ) :
    QMainWindow( parent ),
    fImpl( new Ui::CMainWindow )
{
    auto api = COutlookAPI::getInstance( this );

    fImpl->setupUi( this );

    connect( fImpl->actionSelectServer, &QAction::triggered, this, &CMainWindow::slotSelectServer );

    connect( fImpl->actionReloadAllData, &QAction::triggered, this, &CMainWindow::slotReloadAll );
    connect( fImpl->actionReloadEmail, &QAction::triggered, this, &CMainWindow::slotReloadEmail );
    connect( fImpl->actionReloadFolders, &QAction::triggered, this, &CMainWindow::slotReloadFolders );
    connect( fImpl->actionReloadRules, &QAction::triggered, this, &CMainWindow::slotReloadRules );

    connect( fImpl->actionSortRules, &QAction::triggered, this, &CMainWindow::slotSortRules );
    connect( fImpl->actionRenameRules, &QAction::triggered, this, &CMainWindow::slotRenameRules );
    connect( fImpl->actionMergeRules, &QAction::triggered, this, &CMainWindow::slotMergeRules );
    connect( fImpl->actionEnableAllRules, &QAction::triggered, this, &CMainWindow::slotEnableAllRules );
    connect( fImpl->actionMoveFromToAddress, &QAction::triggered, this, &CMainWindow::slotMoveFromToAddress );

    connect( fImpl->actionAddFolderForSelectedEmail, &QAction::triggered, this, &CMainWindow::slotAddFolderForSelectedEmail );

    connect( fImpl->actionAddRule, &QAction::triggered, this, &CMainWindow::slotAddRule );
    connect( fImpl->actionRunSelectedRule, &QAction::triggered, this, &CMainWindow::slotRunSelectedRule );
    connect( fImpl->actionRunAllRules, &QAction::triggered, this, &CMainWindow::slotRunAllRules );
    connect( fImpl->actionAddToSelectedRule, &QAction::triggered, this, &CMainWindow::slotAddToSelectedRule );

    connect( fImpl->actionProcessAllEmailWhenLessThan200Emails, &QAction::triggered, [ = ]() { api->setProcessAllEmailWhenLessThan200Emails( fImpl->actionProcessAllEmailWhenLessThan200Emails->isChecked() ); } );
    connect( fImpl->actionOnlyProcessUnread, &QAction::triggered, [ = ]() { api->setOnlyProcessUnread( fImpl->actionOnlyProcessUnread->isChecked() ); } );
    connect( fImpl->actionLoadEmailFromJunkFolder, &QAction::triggered, [ = ]() { api->setLoadEmailFromJunkFolder( fImpl->actionLoadEmailFromJunkFolder->isChecked() ); } );

    connect( COutlookAPI::getInstance().get(), &COutlookAPI::sigOptionChanged, this, &CMainWindow::updateWindowTitle );

    connect( fImpl->folders, &CFoldersView::sigFolderSelected, this, &CMainWindow::slotUpdateActions );
    connect( fImpl->email, &CEmailView::sigRuleSelected, this, &CMainWindow::slotUpdateActions );
    connect( fImpl->rules, &CRulesView::sigRuleSelected, this, &CMainWindow::slotUpdateActions );

    setupStatusBar();

    setWindowTitle( QObject::tr( "Rules Maker" ) );

    connect(
        api.get(), &COutlookAPI::sigAccountChanged,
        [ = ]()
        {
            slotUpdateActions();
            slotReloadAll();
        } );

    connect( api.get(), &COutlookAPI::sigInitStatus, this, &CMainWindow::slotInitStatus );
    connect( api.get(), &COutlookAPI::sigSetStatus, this, &CMainWindow::slotSetStatus );
    connect( api.get(), &COutlookAPI::sigIncStatusValue, this, &CMainWindow::slotIncStatusValue );
    connect( api.get(), &COutlookAPI::sigStatusMessage, this, &CMainWindow::slotStatusMessage );
    connect( api.get(), &COutlookAPI::sigStatusFinished, this, &CMainWindow::slotFinishedStatus );

    slotUpdateActions();

    fImpl->actionProcessAllEmailWhenLessThan200Emails->setChecked( api->processAllEmailWhenLessThan200Emails() );
    fImpl->actionOnlyProcessUnread->setChecked( api->onlyProcessUnread() );
    fImpl->actionLoadEmailFromJunkFolder->setChecked( api->loadEmailFromJunkFolder() );

    updateWindowTitle();

    auto updateIcon = [ = ]( QAction *action )
    {
        auto tb = dynamic_cast< QToolButton * >( fImpl->toolBar->widgetForAction( action ) );
        if ( !tb )
            return;
        tb->setToolButtonStyle( Qt::ToolButtonTextBesideIcon );
    };
    updateIcon( fImpl->actionReloadAllData );
    updateIcon( fImpl->actionReloadEmail );
    updateIcon( fImpl->actionReloadFolders );
    updateIcon( fImpl->actionReloadRules );

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
    fImpl->actionRunAllRules->setEnabled( accountSelected );

    auto emailSelected = !fImpl->email->getRulesForSelection().isEmpty();
    auto displayNameOfEmailSelected = fImpl->email->getSelectedDisplayName();
    auto ruleSelected = fImpl->rules->ruleSelected();
    auto folderSelected = !fImpl->folders->selectedPath().isEmpty();

    fImpl->actionAddFolderForSelectedEmail->setEnabled( !displayNameOfEmailSelected.isEmpty() );

    fImpl->actionRunSelectedRule->setEnabled( ruleSelected );
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

void CMainWindow::clearSelection()
{
    fImpl->folders->clearSelection();
    fImpl->email->clearSelection();
    fImpl->rules->clearSelection();
    slotUpdateActions();
}

void CMainWindow::slotAddFolderForSelectedEmail()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );

    auto folderName = fImpl->email->getSelectedDisplayName();
    fImpl->folders->addFolder( folderName );

    qApp->restoreOverrideCursor();
}

void CMainWindow::slotAddRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto destFolder = fImpl->folders->selectedFolder();
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !COutlookAPI::getInstance()->addRule( destFolder, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    clearSelection();
    slotReloadEmail();
    //slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotAddToSelectedRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto rule = fImpl->rules->selectedRule();
    auto rules = fImpl->email->getRulesForSelection();

    QStringList msgs;
    if ( !COutlookAPI::getInstance()->addToRule( rule, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not modify rule\n" + msgs.join( "\n" ) );
    }
    clearSelection();
    slotReloadEmail();
    slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotMergeRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::getInstance()->mergeRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRenameRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::getInstance()->renameRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotSortRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::getInstance()->sortRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotMoveFromToAddress()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::getInstance()->moveFromToAddress() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotEnableAllRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::getInstance()->enableAllRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunAllRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::getInstance()->runAllRules();
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunSelectedRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    fImpl->rules->runSelectedRule();
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotReloadAll()
{
    clearViews();
    updateWindowTitle();
    if ( COutlookAPI::getInstance()->accountSelected() )
    {
        fImpl->folders->reload( true );
        fImpl->rules->reload( true );
        fImpl->email->reload( true );
    }

    slotUpdateActions();
}

void CMainWindow::updateWindowTitle()
{
    auto windowTitle = tr( "Outlook Rules Maker" );
    if ( COutlookAPI::getInstance()->accountSelected() )
    {
        windowTitle += tr( " - %1" ).arg( COutlookAPI::getInstance()->accountName() );
        windowTitle += tr( " - %1" ).arg( COutlookAPI::getInstance()->rootProcessFolderName() );
    }
    setWindowTitle( windowTitle );
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
    auto account = COutlookAPI::getInstance()->selectAccount( false );
    if ( !account )
        return;

    updateWindowTitle();
    clearViews();
    slotReloadAll();
}

bool CMainWindow::running() const
{
    bool running = false;
    for ( auto &&ii : fProgressBars )
    {
        if ( ii.second->isVisible() )
        {
            running = true;
            break;
        }
    }
    return running;
}

void CMainWindow::slotHandleProgressToggle()
{
    bool running = this->running();

    fCancelButton->setVisible( running );
    if ( !running )
    {
        statusBar()->showMessage( QString() );
    }

    auto actions = this->findChildren< QAction * >();
    for ( auto &&action : actions )
    {
        if ( !action->menu() )
            action->setEnabled( !running );
    }
    auto filters = this->findChildren< QLineEdit * >();
    for ( auto &&filter : filters )
    {
        filter->setEnabled( !running );
    }

    auto buttons = this->findChildren< QAbstractButton * >();
    for ( auto &&button : buttons )
    {
        if ( button == fCancelButton )
            continue;
        button->setEnabled( !running );
    }

    if ( !running )
        slotUpdateActions();
}

CStatusProgress *CMainWindow::addStatusBar( QString label, CWidgetWithStatus *object )
{
    Q_ASSERT( ( !label.isEmpty() && !object ) || ( label.isEmpty() && object ) );
    if ( object )
    {
        connect( object, &CWidgetWithStatus::sigStatusMessage, this, &CMainWindow::slotStatusMessage );
        connect( object, &CWidgetWithStatus::sigInitStatus, this, &CMainWindow::slotInitStatus );
        connect( object, &CWidgetWithStatus::sigSetStatus, this, &CMainWindow::slotSetStatus );
        connect( object, &CWidgetWithStatus::sigIncStatusValue, this, &CMainWindow::slotIncStatusValue );
        label = object->statusLabel();
    }

    auto progress = new CStatusProgress( label );
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

    addStatusBar( {}, fImpl->folders );
    addStatusBar( {}, fImpl->email );
    addStatusBar( {}, fImpl->rules );

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

void CMainWindow::slotStatusMessage( const QString &msg )
{
    statusBar()->showMessage( msg, 5000 );
    qApp->processEvents();
}

CStatusProgress *CMainWindow::getProgressBar( const QString &label )
{
    CStatusProgress *bar = nullptr;
    auto pos = fProgressBars.find( label );
    if ( pos == fProgressBars.end() )
        bar = addStatusBar( label, nullptr );
    else
        bar = ( *pos ).second;
    return bar;
}

void CMainWindow::slotSetStatus( const QString &label, int curr, int max )
{
    auto bar = getProgressBar( label );
    if ( !bar )
        return;
    bar->slotSetStatus( curr, max );
}

void CMainWindow::slotInitStatus( const QString &label, int max )
{
    auto bar = getProgressBar( label );
    if ( !bar )
        return;
    bar->setRange( 0, max );
}

void CMainWindow::slotFinishedStatus( const QString &label )
{
    auto bar = getProgressBar( label );
    if ( !bar )
        return;
    bar->finished();
}

void CMainWindow::slotIncStatusValue( const QString &label )
{
    auto bar = getProgressBar( label );
    if ( !bar )
        return;
    bar->slotIncValue();
}
