#include "FoldersView.h"
#include "FoldersModel.h"
#include "ui_FoldersView.h"

#include <QTimer>

CFoldersView::CFoldersView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CFoldersView )
{
    fImpl->setupUi( this );

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload(); } );

    setWindowTitle( QObject::tr( "Folders" ) );
}

CFoldersView::~CFoldersView()
{
}

void CFoldersView::reload()
{
    fModel = std::make_shared< CFoldersModel >( this );
    fImpl->folders->setModel( fModel.get() );
    connect( fImpl->folders->selectionModel(), &QItemSelectionModel::currentChanged, this, &CFoldersView::itemSelected );
    connect(
        fModel.get(), &CFoldersModel::sigFinishedLoading,
        [ = ]()
        {
            fImpl->folders->expandAll();
            emit sigFinishedLoading();
        } );
    fModel->reload();
}

void CFoldersView::itemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;
    auto fullPath = fModel->fullPath( index );
    fImpl->name->setText( fullPath );
}
