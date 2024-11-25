#include "OutlookSetup.h"
#include "FoldersDlg.h"
#include "OutlookAPI.h"

#include "ui_OutlookSetup.h"

#include <QPushButton>
#include <QTimer>
#include "MSOUTL.h"

COutlookSetup::COutlookSetup( QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::COutlookSetup )
{
    fImpl->setupUi( this );
    connect( fImpl->accountBtn, &QToolButton::clicked, this, &COutlookSetup::slotSelectAccount );
    connect( fImpl->folderBtn, &QToolButton::clicked, this, &COutlookSetup::slotSelectFolder );

    setWindowTitle( QObject::tr( "Setup" ) );

    fImpl->folderBtn->setEnabled( true );
    QTimer::singleShot( 0, [ = ]() { slotSelectAccount( true ); } );
}

COutlookSetup::~COutlookSetup()
{
}

void COutlookSetup::slotSelectAccount( bool useInbox )
{
    auto account = COutlookAPI::getInstance()->selectAccount( false, dynamic_cast< QWidget * >( parent() ) );
    if ( !account )
        return;
    fImpl->account->setText( account->DisplayName() );
    fImpl->rootFolder->clear();
    slotSelectFolder( useInbox );
}

void COutlookSetup::slotSelectFolder( bool useInbox )
{
    auto folder = COutlookAPI::getInstance()->selectInbox( dynamic_cast< QWidget * >( parent() ), false ).first;
    if ( !folder )
        return;

    if ( useInbox )
    {
        fImpl->rootFolder->setText( folder->FullFolderPath() );
        COutlookAPI::getInstance()->setRootFolder( folder );
        return;
    }
    CFoldersDlg dlg( this );
    if ( dlg.exec() == QDialog::Accepted )
    {
        fImpl->rootFolder->setText( dlg.fullPath() );
        COutlookAPI::getInstance()->setRootFolder( dlg.selectedFolder() );
    }
}

void COutlookSetup::reject()
{
    COutlookAPI::getInstance()->logout( false );
    QDialog::reject();
}
