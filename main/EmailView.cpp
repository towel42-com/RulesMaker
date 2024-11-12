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

    QTimer::singleShot(
        0,
        [ = ]()
        {
            fModel = std::make_shared< CEmailModel >( this );
            fImpl->emails->setModel( fModel.get() );
            connect( fImpl->emails->selectionModel(), &QItemSelectionModel::currentChanged, this, &CEmailView::itemSelected );
        } );

    connect( fImpl->groupEmails, &QPushButton::clicked, this, &CEmailView::slotLoadGrouped );
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

void CEmailView::slotLoadGrouped()
{
    auto &&[ from, to, cc, subjects ] = fModel->getGroupedEmailModels( this );
    fImpl->fromGroupings->setModel( from );
    fImpl->toGroupings->setModel( to );
    fImpl->ccGroupings->setModel( cc );
    fImpl->subjectGroupings->setModel( subjects );
    connect(
        fModel.get(), &CEmailModel::sigFinishedGroupingEmails,
        [ = ]()
        {
            fImpl->fromGroupings->expandAll();
            fImpl->toGroupings->expandAll();
            fImpl->ccGroupings->expandAll();
            fImpl->subjectGroupings->expandAll();
        } );
}
