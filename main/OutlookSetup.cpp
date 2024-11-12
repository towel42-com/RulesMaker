#include "OutlookSetup.h"
#include "ui_OutlookSetup.h"
#include "OutlookHelpers.h"

#include <QPushButton>
#include "MSOUTL.h"

COutlookSetup::COutlookSetup( QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::COutlookSetup )
{
    fImpl->setupUi( this );
    connect( fImpl->accountBtn, &QPushButton::clicked, this, &COutlookSetup::slotSelectAccount );
    connect( fImpl->inboxBtn, &QPushButton::clicked, this, &COutlookSetup::slotSelectInbox );
    connect( fImpl->contactsBtn, &QPushButton::clicked, this, &COutlookSetup::slotSelectContacts );

    setWindowTitle( QObject::tr( "Setup" ) );

    fImpl->inboxBtn->setEnabled( false );
    fImpl->contactsBtn->setEnabled( false );
}

COutlookSetup::~COutlookSetup()
{
}

void COutlookSetup::slotSelectAccount()
{
    auto account = COutlookHelpers::getInstance()->selectAccount( dynamic_cast< QWidget * >( parent() ) );
    if ( !account )
        return;
    fImpl->account->setText( account->DisplayName() );
    fImpl->inbox->clear();
    fImpl->contacts->clear();
    selectInbox( true );
    selectContacts( true );
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

void COutlookSetup::slotSelectContacts()
{
    selectContacts( false );
}

void COutlookSetup::selectContacts( bool singleOnly )
{
    auto && [ folder, hadMultiple ] = COutlookHelpers::getInstance()->selectContacts( dynamic_cast< QWidget * >( parent() ), singleOnly );
    fImpl->contactsBtn->setEnabled( hadMultiple );

    if ( !folder )
        return;

    fImpl->contacts->setText( folder->FullFolderPath() );
}
