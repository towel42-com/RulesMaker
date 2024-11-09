#include "ContactsView.h"
#include "ContactsModel.h"
#include "ui_ContactsView.h"

#include <QTimer>

CContactsView::CContactsView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CContactsView )
{
    fImpl->setupUi( this );
    connect( fImpl->addButton, &QPushButton::clicked, this, &CContactsView::addEntry );
    connect( fImpl->changeButton, &QPushButton::clicked, this, &CContactsView::changeEntry );

    QTimer::singleShot(
        0,
        [ = ]()
        {
            fModel = std::make_shared< CContactsModel >( this );
            fImpl->contacts->setModel( fModel.get() );
            connect( fImpl->contacts->selectionModel(), &QItemSelectionModel::currentChanged, this, &CContactsView::itemSelected );
        } );

    setWindowTitle( QObject::tr( "Contacts" ) );
}

CContactsView::~CContactsView()
{
}

void CContactsView::updateOutlook()
{
    fModel->update();
}

void CContactsView::addEntry()
{
    if ( !fImpl->firstName->text().isEmpty() || !fImpl->lastName->text().isEmpty() || !fImpl->address->text().isEmpty() || !fImpl->email->text().isEmpty() )
    {
        fModel->addItem( fImpl->firstName->text(), fImpl->lastName->text(), fImpl->address->text(), fImpl->email->text() );
    }

    fImpl->firstName->clear();
    fImpl->lastName->clear();
    fImpl->address->clear();
    fImpl->email->clear();
}

void CContactsView::changeEntry()
{
    QModelIndex current = fImpl->contacts->currentIndex();

    if ( current.isValid() )
        fModel->changeItem( current, fImpl->firstName->text(), fImpl->lastName->text(), fImpl->address->text(), fImpl->email->text() );
}

void CContactsView::itemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;

    QAbstractItemModel *model = fImpl->contacts->model();
    fImpl->firstName->setText( model->data( model->index( index.row(), 0 ) ).toString() );
    fImpl->lastName->setText( model->data( model->index( index.row(), 1 ) ).toString() );
    fImpl->address->setText( model->data( model->index( index.row(), 2 ) ).toString() );
    fImpl->email->setText( model->data( model->index( index.row(), 3 ) ).toString() );
}
