#include "OutlookSetup.h"
#include "ui_OutlookSetup.h"
#include "OutlookHelpers.h"

#include <QPushButton>
#include <QTimer>
#include "MSOUTL.h"

COutlookSetup::COutlookSetup( QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::COutlookSetup )
{
    fImpl->setupUi( this );
    connect( fImpl->accountBtn, &QPushButton::clicked, this, &COutlookSetup::slotSelectAccount );
    connect( fImpl->inboxBtn, &QPushButton::clicked, this, &COutlookSetup::slotSelectInbox );

    setWindowTitle( QObject::tr( "Setup" ) );

    fImpl->inboxBtn->setEnabled( false );
    QTimer::singleShot( 0, [ = ]() { slotSelectAccount(); } );
}

COutlookSetup::~COutlookSetup()
{
}

void COutlookSetup::slotSelectAccount()
{
    auto account = COutlookHelpers::getInstance()->selectAccount( false, dynamic_cast< QWidget * >( parent() ) );
    if ( !account )
        return;
    fImpl->account->setText( account->DisplayName() );
    fImpl->inbox->clear();
    selectInbox( true );
}

void COutlookSetup::slotSelectInbox()
{
    selectInbox( false );
}

void COutlookSetup::selectInbox( bool singleOnly )
{
    auto &&[ folder, hadMultiple ] = COutlookHelpers::getInstance()->selectInbox( dynamic_cast< QWidget * >( parent() ), singleOnly );
    fImpl->inboxBtn->setEnabled( hadMultiple );

    if ( !folder )
        return;

    fImpl->inbox->setText( folder->FullFolderPath() );
}

void COutlookSetup::reject()
{
    COutlookHelpers::getInstance()->logout( false );
    QDialog::reject();
}
