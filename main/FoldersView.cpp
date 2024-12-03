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
        QTimer::singleShot( 0, [ = ]() { reload( true ); } );
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
            fImpl->folders->resizeColumnToContents( 0 );
            fImpl->folders->collapseAll();
            fImpl->folders->expand( fImpl->folders->model()->index( 0, 0 ) );
            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
        } );
    connect( fImpl->addFolder, &QPushButton::clicked, this, &CFoldersView::slotAddFolder );
    connect( fModel.get(), &CFoldersModel::sigIncStatusValue, this, &CFoldersView::sigIncStatusValue );
    connect( fModel.get(), &CFoldersModel::sigSetStatus, this, &CFoldersView::sigSetStatus );
}

CFoldersView::~CFoldersView()
{
}

void CFoldersView::reload( bool notifyOnFinish )
{
    fNotifyOnFinish = notifyOnFinish;
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

    auto currentPath = fModel->pathForItem( index );
    fImpl->name->setText( currentPath );
    emit sigFolderSelected( currentPath );
}

void CFoldersView::slotAddFolder()
{
    auto idx = fImpl->folders->currentIndex();
    fModel->addFolder( idx, this );
    //reload( false );
}

QString CFoldersView::selectedPath() const
{
    auto idx = fImpl->folders->currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->pathForItem( idx );
}

QString CFoldersView::selectedFullPath() const
{
    auto idx = fImpl->folders->currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->fullPathForItem( idx );
}

std::shared_ptr< Outlook::Folder > CFoldersView::selectedFolder() const
{
    auto idx = fImpl->folders->currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->folderForItem( idx );
}
