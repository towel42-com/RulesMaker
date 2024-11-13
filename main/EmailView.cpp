#include "EmailView.h"
#include "EmailModel.h"
#include "EmailGroupingModel.h"

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

    fModel = std::make_shared< CEmailModel >( this );
    fImpl->emails->setModel( fModel.get() );

    fGroupedModel = fModel->getGroupedEmailModel();
    fImpl->fromGroupings->setModel( fGroupedModel );

    connect( fImpl->fromGroupings, &QTreeView::doubleClicked, this, &CEmailView::slotGroupedItemDoubleClicked );
    connect( fImpl->emails->selectionModel(), &QItemSelectionModel::currentChanged, this, &CEmailView::slotItemSelected );
    connect(
        fModel.get(), &CEmailModel::sigFinishedLoading,
        [ = ]()
        {
            slotLoadGrouped();
            emit sigFinishedLoading();
        } );
    setWindowTitle( QObject::tr( "Inbox Emails" ) );
}

CEmailView::~CEmailView()
{
}

void CEmailView::clear()
{
    if ( fModel )
        fModel->clear();
}

void CEmailView::slotItemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;

    QAbstractItemModel *model = fImpl->emails->model();
    fImpl->from->setText( model->data( model->index( index.row(), 0 ) ).toString() );
    fImpl->to->setText( model->data( model->index( index.row(), 1 ) ).toString() );
    fImpl->cc->setText( model->data( model->index( index.row(), 2 ) ).toString() );
    fImpl->subject->setText( model->data( model->index( index.row(), 3 ) ).toString() );
}

void CEmailView::slotGroupedItemDoubleClicked( const QModelIndex &idx )
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
    fModel->reload();
}

void CEmailView::slotLoadGrouped()
{
    connect(
        fModel.get(), &CEmailModel::sigFinishedGrouping,
        [ = ]()
        {
            fImpl->fromGroupings->expandAll();
            emit sigFinishedGrouping();
        } );
}
