#include "FoldersView.h"
#include "FoldersModel.h"
#include "ui_FoldersView.h"

#include <QTimer>

CFoldersView::CFoldersView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CFoldersView )
{
    fImpl->setupUi( this );
    connect( fImpl->addButton, &QPushButton::clicked, this, &CFoldersView::addEntry );
    connect( fImpl->changeButton, &QPushButton::clicked, this, &CFoldersView::changeEntry );

    QTimer::singleShot(
        0,
        [ = ]()
        {
            fModel = std::make_shared< CFoldersModel >( this );
            fImpl->folders->setModel( fModel.get() );
            fImpl->folders->expandAll();
            connect( fImpl->folders->selectionModel(), &QItemSelectionModel::currentChanged, this, &CFoldersView::itemSelected );
        } );

    setWindowTitle( QObject::tr( "Folders" ) );
}

CFoldersView::~CFoldersView()
{
}

void CFoldersView::updateOutlook()
{
    fModel->update();
}

void CFoldersView::addEntry()
{
    if ( !fImpl->name->text().isEmpty() )
    {
        fModel->addItem( fImpl->name->text() );
    }

    fImpl->name->clear();
}

void CFoldersView::changeEntry()
{
    QModelIndex current = fImpl->folders->currentIndex();

    if ( current.isValid() )
        fModel->changeItem( current, fImpl->name->text() );
}

void CFoldersView::itemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;
    auto fullPath = fModel->fullPath( index );
    fImpl->name->setText( fullPath );
}
