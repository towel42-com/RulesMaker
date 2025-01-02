#include "MainWindow.h"
#include "OutlookAPI/OutlookAPI.h"
#include "StatusProgress.h"
#include "Version.h"
#include "Settings.h"

#include "ui_MainWindow.h"

#include <QTimer>
#include <QMessageBox>
#include <QPushButton>
#include <QCursor>
#include <QApplication>
#include <QLineEdit>
#include <QToolButton>
#include <QAbstractItemView>

CMainWindow::CMainWindow( QWidget *parent ) :
    QMainWindow( parent ),
    fImpl( new Ui::CMainWindow )
{
    auto api = COutlookAPI::instance( this );

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
    connect( fImpl->actionAddToSelectedRule, &QAction::triggered, this, &CMainWindow::slotAddToSelectedRule );

    connect( fImpl->actionRunAllRules, &QAction::triggered, this, &CMainWindow::slotRunAllRules );
    connect( fImpl->actionRunAllRulesOnAllFolders, &QAction::triggered, this, &CMainWindow::slotRunAllRulesOnAllFolders );
    connect( fImpl->actionRunSelectedRule, &QAction::triggered, this, &CMainWindow::slotRunSelectedRule );
    connect( fImpl->actionRunAllRulesOnSelectedFolder, &QAction::triggered, this, &CMainWindow::slotRunAllRulesOnSelectedFolder );
    connect( fImpl->actionRunSelectedRuleOnSelectedFolder, &QAction::triggered, this, &CMainWindow::slotRunSelectedRuleOnSelectedFolder );
    connect( fImpl->actionEmptyTrash, &QAction::triggered, this, &CMainWindow::slotEmptyTrash );
    connect( fImpl->actionEmptyJunkFolder, &QAction::triggered, this, &CMainWindow::slotEmptyJunkFolder );

    connect( fImpl->actionSettings, &QAction::triggered, this, &CMainWindow::slotSettings );
    connect( fImpl->actionProcessAllEmailWhenLessThan200Emails, &QAction::triggered, [ = ]() { api->setProcessAllEmailWhenLessThan200Emails( fImpl->actionProcessAllEmailWhenLessThan200Emails->isChecked() ); } );
    connect( fImpl->actionOnlyProcessTheFirst500Emails, &QAction::triggered, [ = ]() { api->setOnlyProcessTheFirst500Emails( fImpl->actionOnlyProcessTheFirst500Emails->isChecked() ); } );

    connect( fImpl->actionOnlyProcessUnread, &QAction::triggered, [ = ]() { api->setOnlyProcessUnread( fImpl->actionOnlyProcessUnread->isChecked() ); } );
    connect( fImpl->actionIncludeJunkFolderWhenRunningOnAllFolders, &QAction::triggered, [ = ]() { api->setIncludeJunkFolderWhenRunningOnAllFolders( fImpl->actionIncludeJunkFolderWhenRunningOnAllFolders->isChecked() ); } );
    connect( fImpl->actionIncludeDeletedFolderWhenRunningOnAllFolders, &QAction::triggered, [ = ]() { api->setIncludeDeletedFolderWhenRunningOnAllFolders( fImpl->actionIncludeDeletedFolderWhenRunningOnAllFolders->isChecked() ); } );
    connect( fImpl->actionDisableRatherThanDeleteRules, &QAction::triggered, [ = ]() { api->setDisableRatherThanDeleteRules( fImpl->actionDisableRatherThanDeleteRules->isChecked() ); } );

    connect( COutlookAPI::instance().get(), &COutlookAPI::sigOptionChanged, this, &CMainWindow::slotOptionsChanged );

    connect(
        fImpl->folders, &CFoldersView::sigFolderSelected,
        [ = ]( const QString &path )
        {
            slotStatusMessage( QString( "Folder Selected: %1" ).arg( path ) );
            slotUpdateActions();
        } );
    connect(
        fImpl->email, &CEmailView::sigEmailSelected,
        [ = ]()
        {
            auto path = fImpl->email->getDisplayTextForSelection();
            slotStatusMessage( QString( "Email Selected: %1" ).arg( path ) );
            slotUpdateActions();
        } );
    connect(
        fImpl->rules, &CRulesView::sigRuleSelected,
        [ = ]()
        {
            auto rule = fImpl->rules->selectedRule();
            auto path = COutlookAPI::instance()->ruleNameForRule( rule, false, true );
            slotStatusMessage( QString( "Rule Selected: %1" ).arg( path ) );
            slotUpdateActions();
        } );

    connect( this, &CMainWindow::sigRunningStateChanged, fImpl->rules, &CRulesView::slotRunningStateChanged );
    connect( this, &CMainWindow::sigRunningStateChanged, fImpl->email, &CEmailView::slotRunningStateChanged );
    connect( this, &CMainWindow::sigRunningStateChanged, fImpl->folders, &CFoldersView::slotRunningStateChanged );

    connect( fImpl->actionAbout, &QAction::triggered, this, &CMainWindow::slotAbout );
    setupStatusBar();

    setWindowTitle( NVersion::APP_NAME );

    connect(
        api.get(), &COutlookAPI::sigAccountChanged,
        [ = ]()
        {
            updateActions();
            slotReloadAll();
        } );

    connect( api.get(), &COutlookAPI::sigInitStatus, this, &CMainWindow::slotInitStatus );
    connect( api.get(), &COutlookAPI::sigSetStatus, this, &CMainWindow::slotSetStatus );
    connect( api.get(), &COutlookAPI::sigIncStatusValue, this, &CMainWindow::slotIncStatusValue );
    connect( api.get(), &COutlookAPI::sigStatusMessage, this, &CMainWindow::slotStatusMessage );
    connect( api.get(), &COutlookAPI::sigStatusFinished, this, &CMainWindow::slotFinishedStatus );

    updateActions();

    slotOptionsChanged();

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
    COutlookAPI::instance()->logout( false );
}

void CMainWindow::slotUpdateActions()
{
    updateActions();
}

void CMainWindow::updateActions()
{
    TReason accountSelected( COutlookAPI::instance()->accountSelected(), "Rule not selected" );
    TReason emailSelected( !fImpl->email->getDisplayTextForSelection().isEmpty(), "Email not selected" );
    TReason emailHasDisplayName( !fImpl->email->getEmailDisplayNameForSelection().isEmpty(), "Selected email does not have a display name" );
    TReason ruleSelected( fImpl->rules->ruleSelected(), "Rule not selected" );
    TReason folderSelected( !fImpl->folders->selectedPath().isEmpty(), "Folder not selected" );
    TReason folderSame( true, "Selected folder does not match selected rule's target folder" );

    setEnabled( fImpl->actionReloadAllData, accountSelected );
    setEnabled( fImpl->actionReloadEmail, accountSelected );
    setEnabled( fImpl->actionReloadFolders, accountSelected );
    setEnabled( fImpl->actionReloadRules, accountSelected );
    setEnabled( fImpl->actionSortRules, accountSelected );
    setEnabled( fImpl->actionRenameRules, accountSelected );
    setEnabled( fImpl->actionMoveFromToAddress, accountSelected );
    setEnabled( fImpl->actionReloadAllData, accountSelected );
    setEnabled( fImpl->actionRunAllRules, accountSelected );
    setEnabled( fImpl->actionRunAllRulesOnAllFolders, accountSelected );
    setEnabled( fImpl->actionEmptyTrash, accountSelected );
    setEnabled( fImpl->actionEmptyJunkFolder, accountSelected );

    if ( emailSelected.first && ruleSelected.first )
    {
        auto selectedFolder = fImpl->folders->selectedFolder();
        if ( selectedFolder )
        {
            auto ruleFolder = fImpl->rules->folderForSelectedRule();
            auto selectedFolderPath = fImpl->folders->selectedFullPath();
            folderSame.first = ruleFolder == selectedFolderPath;
        }
        else
            folderSame.first = true;
    }

    if ( emailSelected.first && !emailHasDisplayName.first )
        setEnabled( fImpl->actionAddFolderForSelectedEmail, emailHasDisplayName );
    else
        setEnabled( fImpl->actionAddFolderForSelectedEmail, emailSelected );

    setEnabled( fImpl->actionRunSelectedRule, ruleSelected );
    setEnabled( fImpl->actionAddToSelectedRule, { emailSelected, ruleSelected, folderSame } );

    setEnabled( fImpl->actionRunAllRulesOnSelectedFolder, folderSelected );
    setEnabled( fImpl->actionRunSelectedRuleOnSelectedFolder, { folderSelected, ruleSelected } );

    setEnabled( fImpl->actionAddRule, { accountSelected, folderSelected, emailSelected } );
}

void CMainWindow::clearSelection()
{
    fImpl->folders->clearSelection();
    fImpl->email->clearSelection();
    fImpl->rules->clearSelection();
    updateActions();
}

void CMainWindow::slotAddFolderForSelectedEmail()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );

    auto folderName = fImpl->email->getEmailDisplayNameForSelection();
    fImpl->folders->addFolder( folderName );

    qApp->restoreOverrideCursor();
}

void CMainWindow::slotAddRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto destFolder = fImpl->folders->selectedFolder();
    auto rules = fImpl->email->getMatchTextForSelection();

    QStringList msgs;
    if ( !COutlookAPI::instance()->addRule( destFolder, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
    }
    clearSelection();
    slotReloadEmail();
    slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotAddToSelectedRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto rule = fImpl->rules->selectedRule();
    auto rules = fImpl->email->getMatchTextForSelection();

    QStringList msgs;
    if ( !COutlookAPI::instance()->addToRule( rule, rules, msgs ) )
    {
        QMessageBox::critical( this, "Error", "Could not modify rule\n" + msgs.join( "\n" ) );
    }
    clearSelection();
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotMergeRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::instance()->mergeRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRenameRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::instance()->renameRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotSortRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::instance()->sortRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotMoveFromToAddress()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::instance()->moveFromToAddress() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotEnableAllRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    if ( COutlookAPI::instance()->enableAllRules() )
        slotReloadRules();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunSelectedRule()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto selectedRule = fImpl->rules->selectedRule();
    if ( !selectedRule )
        return;

    COutlookAPI::instance()->runRule( selectedRule );

    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunSelectedRuleOnSelectedFolder()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto destFolder = fImpl->folders->selectedFolder();
    if ( !destFolder )
        return;

    auto selectedRule = fImpl->rules->selectedRule();
    if ( !selectedRule )
        return;

    COutlookAPI::instance()->runRule( selectedRule, destFolder );

    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunAllRules()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->runAllRules();
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunAllRulesOnAllFolders()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->runAllRulesOnAllFolders();
    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotRunAllRulesOnSelectedFolder()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    auto destFolder = fImpl->folders->selectedFolder();
    if ( !destFolder )
        return;

    COutlookAPI::instance()->runAllRules( destFolder );

    slotReloadEmail();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotEmptyTrash()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->emptyTrash();
    fImpl->folders->reloadTrash();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotEmptyJunkFolder()
{
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->emptyJunk();
    fImpl->folders->reloadJunk();
    qApp->restoreOverrideCursor();
}

void CMainWindow::slotReloadAll()
{
    clearViews();
    updateWindowTitle();
    if ( COutlookAPI::instance()->accountSelected() )
    {
        fImpl->folders->reload( true );
        fImpl->rules->reload( true );
        fImpl->email->reload( true );
    }

    updateActions();
}

void CMainWindow::updateWindowTitle()
{
    auto windowTitle = NVersion::APP_NAME + " - " + NVersion::getVersionString( true );
    if ( COutlookAPI::instance()->accountSelected() )
    {
        windowTitle += tr( " - %1" ).arg( COutlookAPI::instance()->accountName() );
        windowTitle += tr( " - %1" ).arg( COutlookAPI::instance()->rootFolderName() );
    }
    setWindowTitle( windowTitle );
}

void CMainWindow::slotReloadEmail()
{
    fImpl->email->clear();
    if ( COutlookAPI::instance()->accountSelected() )
        fImpl->email->reload( false );
    updateActions();
}

void CMainWindow::slotReloadFolders()
{
    fImpl->folders->clear();
    if ( COutlookAPI::instance()->accountSelected() )
        fImpl->folders->reload( false );
    updateActions();
}

void CMainWindow::slotReloadRules()
{
    fImpl->rules->clear();
    if ( COutlookAPI::instance()->accountSelected() )
        fImpl->rules->reload( false );
    updateActions();
}

void CMainWindow::clearViews()
{
    fImpl->email->clear();
    fImpl->rules->clear();
    fImpl->folders->clear();
}

void CMainWindow::slotSelectServer()
{
    if ( !COutlookAPI::instance()->closeAndSelectAccount( false ) )
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
    static std::optional< bool > sPrevRunning;

    bool running = this->running();

    fCancelButton->setVisible( running );

    setEnabled< QAction * >();
    setEnabled< QLineEdit * >();
    setEnabled< QAbstractButton * >();
    setEnabled< QAbstractItemView * >();
    setEnabled< QLabel * >();

    updateActions();
    if ( !sPrevRunning.has_value() || sPrevRunning.value() != running )
    {
        emit sigRunningStateChanged( running );
    }
    sPrevRunning = running;
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
    progress->setVisible( false );
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
    connect( fCancelButton, &QPushButton::clicked, COutlookAPI::instance().get(), &COutlookAPI::slotCanceled );
    connect(
        fCancelButton, &QPushButton::clicked, this,
        [ = ]()
        {
            for ( auto &&ii : fProgressBars )
                ii.second->hide();
            fCancelButton->hide();
            slotHandleProgressToggle();
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

void CMainWindow::slotAbout()
{
    auto title = tr( "About %1" ).arg( NVersion::APP_NAME );
    auto caption = tr( "<h3>About %1</h3>"
                       "<p>%1 version %2</p>"
                       "<p>It is an opensource application licensed under the MIT license.</p>" )
                       .arg( NVersion::APP_NAME )
                       .arg( NVersion::getVersionString( true ) );
    auto aboutText = tr( "<p>It is a tool to help create and maintain Outlook Rules to keep your inbox clean.</p>"
                         "<p>It is designed to work with Microsoft Outlook.</p>"
                         "<p>It is provided under the terms of the MIT License.</p>"
                         "<p>For more information, please visit <a href=\"%1\">%1</a>.</p>"
                         R"(<hr style="width:50%;text-align:left;margin-left:0">)"
                         "<p>%2</p>" )
                         .arg( "https://" + NVersion::PRODUCT_HOMEPAGE, NVersion::COPYRIGHT );

    auto msgBox = new QMessageBox( this );
    msgBox->setAttribute( Qt::WA_DeleteOnClose );
    msgBox->setWindowTitle( title );
    msgBox->setText( caption );
    msgBox->setInformativeText( aboutText );
    msgBox->setTextInteractionFlags( Qt::TextBrowserInteraction );
    auto btn = msgBox->addButton( "&About Qt", QMessageBox::ActionRole );
    connect( btn, &QAbstractButton::clicked, [ = ]() { QMessageBox::aboutQt( msgBox ); } );
    msgBox->addButton( QMessageBox::Ok );
    QPixmap pm( QLatin1String( ":resources/app.png" ) );
    if ( !pm.isNull() )
        msgBox->setIconPixmap( pm );
    msgBox->exec();
}

void CMainWindow::slotSettings()
{
    CSettings settings( this );
    if ( ( settings.exec() == QDialog::Accepted ) && settings.changed() )
    {
        slotOptionsChanged();
    }
}

void CMainWindow::slotOptionsChanged()
{
    auto api = COutlookAPI::instance();

    fImpl->actionProcessAllEmailWhenLessThan200Emails->setChecked( api->processAllEmailWhenLessThan200Emails() );
    fImpl->actionOnlyProcessTheFirst500Emails->setChecked( api->onlyProcessTheFirst500Emails() );
    fImpl->actionOnlyProcessUnread->setChecked( api->onlyProcessUnread() );
    fImpl->actionIncludeJunkFolderWhenRunningOnAllFolders->setChecked( api->includeJunkFolderWhenRunningOnAllFolders() );
    fImpl->actionIncludeDeletedFolderWhenRunningOnAllFolders->setChecked( api->includeDeletedFolderWhenRunningOnAllFolders() );
    fImpl->actionDisableRatherThanDeleteRules->setChecked( api->disableRatherThanDeleteRules() );

    updateWindowTitle();
}
