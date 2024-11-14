#include "FoldersView.h"
#include "FoldersModel.h"
#include "ui_FoldersView.h"

#include <QTimer>

CFoldersView::CFoldersView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CFoldersView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload(); } );

}

void CFoldersView::init()
{
    fImpl->setupUi( this );
    setWindowTitle( QObject::tr( "Folders" ) );

    fModel = std::make_shared< CFoldersModel >( this );
    fImpl->folders->setModel( fModel.get() );
    connect( fImpl->folders->selectionModel(), &QItemSelectionModel::currentChanged, this, &CFoldersView::slotItemSelected );
    connect(
        fModel.get(), &CFoldersModel::sigFinishedLoading,
        [ = ]()
        {
            fImpl->folders->expandAll();
            emit sigFinishedLoading();
        } );
    connect( fImpl->addFolder, &QPushButton::clicked, this, &CFoldersView::slotAddFolder );
}

CFoldersView::~CFoldersView()
{
}

void CFoldersView::reload()
{
    fModel->reload();
}

void CFoldersView::clear()
{
    if ( fModel )
        fModel->clear();
}

void CFoldersView::slotItemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
    {
        emit sigFolderSelected( QString() );
        return;
    }

    auto currentPath = fModel->currentPath( index );
    fImpl->name->setText( currentPath );
    emit sigFolderSelected( currentPath );
}

void CFoldersView::slotAddFolder()
{
    auto idx = fImpl->folders->currentIndex();
    fModel->addFolder( idx, this );
}

QString CFoldersView::currentPath() const
{
    auto idx = fImpl->folders->currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->currentPath( idx );
}

QString CFoldersView::fullPath() const
{
    auto idx = fImpl->folders->currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->fullPath( idx );
}
