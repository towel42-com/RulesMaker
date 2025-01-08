#include "MainWindow.h"
#include "ShowRule.h"
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
    auto api = COutlookAPI::instance( [ this ]( std::shared_ptr< Outlook::Rule > rule ) { return showRule( rule ); }, this );

    fImpl->setupUi( this );

    connect( fImpl->actionSelectServer, &QAction::triggered, this, &CMainWindow::slotSelectServer );
    connect( fImpl->actionExit, &QAction::triggered, this, &CMainWindow::close );

    connect( fImpl->actionReloadAllData, &QAction::triggered, this, &CMainWindow::slotReloadAll );
    connect( fImpl->actionReloadEmail, &QAction::triggered, this, &CMainWindow::slotReloadEmail );
    connect( fImpl->actionReloadFolders, &QAction::triggered, this, &CMainWindow::slotReloadFolders );
    connect( fImpl->actionReloadRules, &QAction::triggered, this, &CMainWindow::slotReloadRules );

    connect( fImpl->actionSortRules, &QAction::triggered, this, &CMainWindow::slotSortRules );
    connect( fImpl->actionRenameRules, &QAction::triggered, this, &CMainWindow::slotRenameRules );
    connect( fImpl->actionMergeRules, &QAction::triggered, this, &CMainWindow::slotMergeRules );
    connect( fImpl->actionEnableAllRules, &QAction::triggered, this, &CMainWindow::slotEnableAllRules );
    connect( fImpl->actionMoveFromToAddress, &QAction::triggered, this, &CMainWindow::slotMoveFromToAddress );
    connect( fImpl->actionFixFromMessageHeaderRules, &QAction::triggered, this, &CMainWindow::slotFixFromMessageHeaderRules );

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
            auto path = fImpl->email->getEmailPatternForSelection();
            slotStatusMessage( QString( "Email Selected: %1" ).arg( path ) );
            slotUpdateActions();
        } );
    connect( fImpl->email, &CEmailView::sigFilterTypeChanged, this, &CMainWindow::slotUpdateActions );
    connect(
        fImpl->rules, &CRulesView::sigRuleSelected,
        [ = ]()
        {
            auto rule = fImpl->rules->selectedRule();
            auto path = COutlookAPI::rawRuleNameForRule( rule );
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
    TReason emailSelected( !fImpl->email->getEmailPatternForSelection().isEmpty(), "Email not selected" );
    TReason emailHasDisplayName( !fImpl->email->selectionHasDisplayName(), "Selected email does not have a display name" );
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
    setEnabled( fImpl->actionFixFromMessageHeaderRules, accountSelected );
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
    setWaitCursor( true );

    auto folderName = fImpl->email->getDisplayNameForSingleSelection();
    fImpl->folders->addFolder( folderName );

    setWaitCursor( false );
}

void CMainWindow::slotAddRule()
{
    setWaitCursor( true );
    auto destFolder = fImpl->folders->selectedFolder();
    auto patterns = fImpl->email->getPatternsForSelection();

    QStringList msgs;
    auto aOK = COutlookAPI::instance()->addRule( destFolder, patterns, msgs );

    if ( !aOK.has_value() )
    {
        setWaitCursor( false );
        return;
    }

    if ( !aOK.value() )
    {
        QMessageBox::critical( this, "Error", "Could not create rule\n" + msgs.join( "\n" ) );
        setWaitCursor( false );
    }
    clearSelection();
    slotReloadEmail();
    slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotAddToSelectedRule()
{
    setWaitCursor( true );
    auto rule = fImpl->rules->selectedRule();
    auto patterns = fImpl->email->getPatternsForSelection();

    QStringList msgs;
    auto aOK = COutlookAPI::instance()->addToRule( rule, patterns, msgs, true );
    if ( !aOK.has_value() )
    {
        setWaitCursor( false );
        return;
    }

    if ( !aOK.value() )
    {
        QMessageBox::critical( this, "Error", "Could not modify rule\n" + msgs.join( "\n" ) );
    }
    clearSelection();
    slotReloadEmail();
    slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotMergeRules()
{
    setWaitCursor( true );
    if ( COutlookAPI::instance()->mergeRules() )
        slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotRenameRules()
{
    setWaitCursor( true );
    if ( COutlookAPI::instance()->renameRules() )
        slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotSortRules()
{
    setWaitCursor( true );
    if ( COutlookAPI::instance()->sortRules() )
        slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotMoveFromToAddress()
{
    setWaitCursor( true );
    if ( COutlookAPI::instance()->moveFromToAddress() )
        slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotFixFromMessageHeaderRules()
{
    setWaitCursor( true );
    if ( COutlookAPI::instance()->fixFromMessageHeaderRules() )
        slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotEnableAllRules()
{
    setWaitCursor( true );
    if ( COutlookAPI::instance()->enableAllRules() )
        slotReloadRules();
    setWaitCursor( false );
}

void CMainWindow::slotRunSelectedRule()
{
    setWaitCursor( true );
    auto selectedRule = fImpl->rules->selectedRule();
    if ( !selectedRule )
        return;

    COutlookAPI::instance()->runRule( selectedRule );

    slotReloadEmail();
    setWaitCursor( false );
}

void CMainWindow::slotRunSelectedRuleOnSelectedFolder()
{
    setWaitCursor( true );
    auto destFolder = fImpl->folders->selectedFolder();
    if ( !destFolder )
        return;

    auto selectedRule = fImpl->rules->selectedRule();
    if ( !selectedRule )
        return;

    COutlookAPI::instance()->runRule( selectedRule, destFolder );

    slotReloadEmail();
    setWaitCursor( false );
}

void CMainWindow::slotRunAllRules()
{
    setWaitCursor( true );
    COutlookAPI::instance()->runAllRules();
    slotReloadEmail();
    setWaitCursor( false );
}

void CMainWindow::slotRunAllRulesOnAllFolders()
{
    setWaitCursor( true );
    COutlookAPI::instance()->runAllRulesOnAllFolders();
    slotReloadEmail();
    setWaitCursor( false );
}

void CMainWindow::slotRunAllRulesOnSelectedFolder()
{
    setWaitCursor( true );
    auto destFolder = fImpl->folders->selectedFolder();
    if ( !destFolder )
        return;

    COutlookAPI::instance()->runAllRules( destFolder );

    slotReloadEmail();
    setWaitCursor( false );
}

void CMainWindow::slotEmptyTrash()
{
    setWaitCursor( true );
    COutlookAPI::instance()->emptyTrash();
    fImpl->folders->reloadTrash();
    setWaitCursor( false );
}

void CMainWindow::slotEmptyJunkFolder()
{
    setWaitCursor( true );
    COutlookAPI::instance()->emptyJunk();
    fImpl->folders->reloadJunk();
    setWaitCursor( false );
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

void CMainWindow::setWaitCursor( bool wait )
{
    if ( wait )
        qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    else
        qApp->restoreOverrideCursor();
}

bool CMainWindow::showRule( std::shared_ptr< Outlook::Rule > rule )
{
    return showRuleDialog( rule, true );
}

bool CMainWindow::editRule( std::shared_ptr< Outlook::Rule > rule )
{
    return showRuleDialog( rule, false );
}

bool CMainWindow::showRuleDialog( std::shared_ptr< Outlook::Rule > rule, bool readOnly )
{
    CShowRule ruleDlg( rule, readOnly, this );
    bool restoreOverride = qApp->overrideCursor() != nullptr;
    if ( restoreOverride )
        setWaitCursor( false );

    if ( ruleDlg.exec() != QDialog::Accepted )
    {
        if ( restoreOverride )
            setWaitCursor( true );
        return false;
    }

    if ( restoreOverride )
        setWaitCursor( true );
    return true;
}
