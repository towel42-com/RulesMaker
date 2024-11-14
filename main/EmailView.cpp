#include "EmailView.h"
#include "GroupedEmailModel.h"

#include "ui_EmailView.h"
#include "MSOUTL.h"

#include <QTimer>

CEmailView::CEmailView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CEmailView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload(); } );
}

void CEmailView::init()
{
    fImpl->setupUi( this );

    fGroupedModel = new CGroupedEmailModel( this );
    fImpl->groupedEmails->setModel( fGroupedModel );

    connect( fImpl->groupedEmails, &QTreeView::doubleClicked, this, &CEmailView::slotItemDoubleClicked );
    connect( fImpl->groupedEmails->selectionModel(), &QItemSelectionModel::currentChanged, this, &CEmailView::slotItemSelected );
    connect(
        fGroupedModel, &CGroupedEmailModel::sigFinishedGrouping,
        [ = ]()
        {
            fImpl->groupedEmails->expandAll();
            emit sigFinishedLoading();
        } );
    setWindowTitle( QObject::tr( "Inbox Emails" ) );
}

CEmailView::~CEmailView()
{
}

void CEmailView::clear()
{
    if ( fGroupedModel )
        fGroupedModel->clear();
}

void CEmailView::slotItemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;

    auto rules = fGroupedModel->rulesForIndex( index );
    fImpl->rule->setText( rules.join( " or " ) );
    emit sigRuleSelected();
}

void CEmailView::slotItemDoubleClicked( const QModelIndex &idx )
{
    if ( !idx.isValid() )
        return;
    auto item = fGroupedModel->emailItemFromIndex( idx );
    if ( !item )
        return;
    item->Display();
}

void CEmailView::reload()
{
    fGroupedModel->reload();
}

QStringList CEmailView::currentRule() const
{
    auto idx = fImpl->groupedEmails->currentIndex();
    if ( !idx.isValid() )
        return {};
    return fGroupedModel->rulesForIndex( idx );
}

void CEmailView::setOnlyGroupUnread( bool value )
{
    fGroupedModel->setOnlyGroupUnread( value );
}

bool CEmailView::onlyGroupUnread() const
{
    return fGroupedModel->onlyGroupUnread();
}


