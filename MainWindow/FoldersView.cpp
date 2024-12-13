#include "FoldersView.h"
#include "FoldersModel.h"
#include "ui_FoldersView.h"
#include "OutlookAPI/OutlookAPI.h"
#include "ListFilterModel.h"

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

    fFilterModel = new CListFilterModel( this );
    fFilterModel->setSourceModel( fModel );
    fImpl->folders->setModel( fFilterModel );
    fImpl->setRootFolderBtn->setEnabled( false );
    fImpl->addFolder->setEnabled( false );
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

    connect( fModel, &CFoldersModel::sigFinishedLoadingChildren, [ = ]( QStandardItem * /*parent*/ ) { fFilterModel->sort( 0, Qt::SortOrder::AscendingOrder ); } );

    connect( fImpl->addFolder, &QPushButton::clicked, this, &CFoldersView::slotAddFolder );
    connect( fModel, &CFoldersModel::sigSetStatus, [ = ]( int curr, int max ) 
        { 
            emit sigSetStatus( statusLabel(), curr, max ); 
            if ( ( max > 10 ) && ( curr == 1 ) || ( ( curr % 10 ) == 0 ) )
            {
                fImpl->folders->expand( fImpl->folders->model()->index( 0, 0 ) );
                fImpl->folders->resizeColumnToContents( 0 );
            }
        } );
    connect( fImpl->filter, &QLineEdit::textChanged, fFilterModel, &CListFilterModel::slotSetFilter );
    connect( fImpl->filter, &QLineEdit::textChanged, [ = ]() { fImpl->folders->expandAll(); } );
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

QModelIndex CFoldersView::sourceIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() || ( idx.model() == fModel ) )
        return idx;
    return fFilterModel->mapToSource( idx );
}

QModelIndex CFoldersView::currentIndex() const
{
    auto filterIdx = fImpl->folders->currentIndex();
    if ( !filterIdx.isValid() )
        return filterIdx;
    return sourceIndex( filterIdx );
}

void CFoldersView::slotSetRootFolder()
{
    auto idx = currentIndex();
    if ( !idx.isValid() )
        return;
    auto folder = fModel->folderForItem( idx );
    if ( !folder )
        return;
    COutlookAPI::instance()->setRootFolder( folder );
}

void CFoldersView::slotItemSelected( const QModelIndex &index )
{
    auto path = index.isValid() ? fModel->pathForItem( index ) : QString();

    fImpl->setRootFolderBtn->setEnabled( index.isValid() );
    fImpl->addFolder->setEnabled( index.isValid() );

    emit sigFolderSelected( path );
}

void CFoldersView::addFolder( const QString &folderName )
{
    auto newIndex = fFilterModel->mapFromSource( fModel->addFolder( currentIndex(), folderName ) );
    selectAndScroll( newIndex );
}

void CFoldersView::slotAddFolder()
{
    auto newIndex = fFilterModel->mapFromSource( fModel->addFolder( currentIndex(), this ) );
    selectAndScroll( newIndex );
}

void CFoldersView::selectAndScroll( const QModelIndex &newIndex )
{
    if ( !newIndex.isValid() )
        return;

    fImpl->folders->setCurrentIndex( newIndex );
    fImpl->folders->scrollTo( newIndex );
}

QString CFoldersView::selectedPath() const
{
    auto idx = currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->pathForItem( idx );
}

QString CFoldersView::selectedFullPath() const
{
    auto idx = currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->fullPathForItem( idx );
}

std::shared_ptr< Outlook::Folder > CFoldersView::selectedFolder() const
{
    auto idx = currentIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->folderForItem( idx );
}
