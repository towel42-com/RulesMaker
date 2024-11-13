#include "MainWindow.h"
#include "OutlookHelpers.h"
#include "OutlookSetup.h"

#include "ui_MainWindow.h"

#include <QTimer>

CMainWindow::CMainWindow( QWidget *parent ) :
    QMainWindow( parent ),
    fImpl( new Ui::CMainWindow )
{
    fImpl->setupUi( this );

    connect( fImpl->actionSelectServer, &QAction::triggered, [ = ]() { slotSelectServer(); } );
    connect( fImpl->actionReloadData, &QAction::triggered, [ = ]() { slotReload(); } );

    connect( fImpl->folders, &CFoldersView::sigFinishedLoading, [ = ]() { fImpl->rules->reload(); } );
    connect( fImpl->rules, &CRulesView::sigFinishedLoading, [ = ]() { fImpl->email->reload(); } );

    setWindowTitle( QObject::tr( "Rules Maker" ) );

    connect(
        COutlookHelpers::getInstance().get(), &COutlookHelpers::sigAccountChanged,
        [ = ]()
        {
            slotUpdateActions();
            slotReload();
        } );
    slotUpdateActions();
}

void CMainWindow::slotUpdateActions()
{
    fImpl->actionReloadData->setEnabled( COutlookHelpers::getInstance()->accountSelected() );
    fImpl->actionAddRule->setEnabled( COutlookHelpers::getInstance()->accountSelected() );
}

CMainWindow::~CMainWindow()
{
    clearViews();
    COutlookHelpers::getInstance()->logout( false );
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
    COutlookSetup dlg;
    bool wasLoaded = ( COutlookHelpers::getInstance()->accountSelected() );
    if ( dlg.exec() == QDialog::Accepted )
        slotReload();
}
