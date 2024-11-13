#include "EmailView.h"
#include "EmailModel.h"
#include "EmailGroupingModel.h"

#include "ui_EmailView.h"

#include <QTimer>

CEmailView::CEmailView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CEmailView )
{
    fImpl->setupUi( this );

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload(); } );

    setWindowTitle( QObject::tr( "Inbox Emails" ) );
}

CEmailView::~CEmailView()
{
}

void CEmailView::itemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;

    QAbstractItemModel *model = fImpl->emails->model();
    fImpl->from->setText( model->data( model->index( index.row(), 0 ) ).toString() );
    fImpl->to->setText( model->data( model->index( index.row(), 1 ) ).toString() );
    fImpl->cc->setText( model->data( model->index( index.row(), 2 ) ).toString() );
    fImpl->subject->setText( model->data( model->index( index.row(), 3 ) ).toString() );
}

void CEmailView::reload()
{
    fModel = std::make_shared< CEmailModel >( this );
    fImpl->emails->setModel( fModel.get() );
    connect( fImpl->emails->selectionModel(), &QItemSelectionModel::currentChanged, this, &CEmailView::itemSelected );
    connect(
        fModel.get(), &CEmailModel::sigFinishedLoading,
        [ = ]()
        {
            slotLoadGrouped();
            emit sigFinishedLoading();
        } );
    fModel->reload();
}

void CEmailView::slotLoadGrouped()
{
    auto &&from = fModel->getGroupedEmailModels( this );
    fImpl->fromGroupings->setModel( from );
    connect(
        fModel.get(), &CEmailModel::sigFinishedGrouping,
        [ = ]()
        {
            fImpl->fromGroupings->expandAll();
            emit sigFinishedGrouping();
        } );
}
