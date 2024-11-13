#include "MainWindow.h"
#include "EmailModel.h"
#include "EmailGroupingModel.h"

#include "ui_MainWindow.h"

#include <QTimer>

CMainWindow::CMainWindow( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CMainWindow )
{
    fImpl->setupUi( this );

    connect( fImpl->reloadBtn, &QPushButton::clicked, [ = ]() { slotReload(); } );
    connect( fImpl->folders, &CFoldersView::sigFinishedLoading, [ = ]() { fImpl->rules->reload(); } );
    connect( fImpl->rules, &CRulesView::sigFinishedLoading, [ = ]() { fImpl->email->reload(); } );

    QTimer::singleShot( 0, [ = ]() { slotReload(); } );
    setWindowTitle( QObject::tr( "Rules Maker" ) );
}

CMainWindow::~CMainWindow()
{
}

void CMainWindow::slotReload()
{
    fImpl->folders->reload();
}