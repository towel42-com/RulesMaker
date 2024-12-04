#include "FoldersView.h"
#include "FoldersModel.h"
#include "ui_FoldersView.h"
#include "OutlookAPI.h"

#include <QTimer>

CFoldersView::CFoldersView( QWidget *parent ) :
    CWidgetWithStatus( parent ),
    fImpl( new Ui::CFoldersView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload( true ); } );
}

void CFoldersView::init()
{
    fImpl->setupUi( this );
    setStatusLabel( "Loading Folders:" );

    fModel = new CFoldersModel( this );
    fImpl->folders->setModel( fModel );
    fImpl->setRootFolderBtn->setEnabled( false );
    connect( fImpl->setRootFolderBtn, &QPushButton::clicked, this, &CFoldersView::slotSetRootFolder );
    connect( fImpl->folders->selectionModel(), &QItemSelectionModel::currentChanged, this, &CFoldersView::slotItemSelected );
    connect(
        fModel, &CFoldersModel::sigFinishedLoading,
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
    connect( fModel, &CFoldersModel::sigSetStatus, [ = ]( int curr, int max ) { emit sigSetStatus( statusLabel(), curr, max ); } );
    connect(
        fModel, &CFoldersModel::sigSetStatus,
        [ = ]( int curr, int max )
        {
            if ( ( max > 10 ) && ( curr == 1 ) || ( ( curr % 10 ) == 0 ) )
            {
                fImpl->folders->expand( fImpl->folders->model()->index( 0, 0 ) );
                fImpl->folders->resizeColumnToContents( 0 );
            }
        } );
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

void CFoldersView::clearSelection()
{
    fImpl->folders->clearSelection();
    fImpl->folders->setCurrentIndex( {} );
    slotItemSelected( {} );
}

void CFoldersView::slotSetRootFolder()
{
    auto idx = fImpl->folders->currentIndex();
    if ( !idx.isValid() )
        return;
    auto folder = fModel->folderForItem( idx );
    if ( !folder )
        return;
    COutlookAPI::getInstance()->setRootFolder( folder );
}

void CFoldersView::slotItemSelected( const QModelIndex &index )
{
    auto currentPath = index.isValid() ? fModel->pathForItem( index ) : QString();
    fImpl->setRootFolderBtn->setEnabled( index.isValid() );
    fImpl->name->setText( currentPath );
    emit sigFolderSelected( currentPath );
}

void CFoldersView::slotAddFolder()
{
    auto idx = fImpl->folders->currentIndex();
    fModel->addFolder( idx, this );
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
