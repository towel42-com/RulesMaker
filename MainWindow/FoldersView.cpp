#include "FoldersView.h"
#include "Models/FoldersModel.h"
#include "Models/ListFilterModel.h"

#include "OutlookAPI/OutlookAPI.h"

#include "ui_FoldersView.h"
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
    connect( fImpl->folders->selectionModel(), &QItemSelectionModel::selectionChanged, this, &CFoldersView::slotItemSelected );
    connect(
        fModel, &CFoldersModel::sigFinishedLoading,
        [ = ]()
        {
            resizeToContentZero( fImpl->folders, EExpandMode::eCollapseAll );
            fImpl->folders->expand( fFilterModel->index( 0, 0 ) );
            auto inboxIndex = fFilterModel->mapFromSource( fModel->inboxIndex() );
            if ( inboxIndex.isValid() )
                fImpl->folders->expand( inboxIndex );
            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
            fImpl->summary->setText( fModel->summary() );
        } );

    connect( fModel, &CFoldersModel::sigFinishedLoadingChildren, [ = ]( QStandardItem * /*parent*/ ) { fFilterModel->sort( 0, Qt::SortOrder::AscendingOrder ); } );

    connect( fImpl->addFolder, &QPushButton::clicked, this, &CFoldersView::slotAddFolder );
    connect(
        fModel, &CFoldersModel::sigSetStatus,
        [ = ]( int curr, int max )
        {
            emit sigSetStatus( statusLabel(), curr, max );
            if ( ( max > 10 ) && ( curr == 1 ) || ( ( curr % 10 ) == 0 ) )
            {
                fImpl->folders->expand( fFilterModel->index( 0, 0 ) );
                auto inboxIndex = fFilterModel->mapFromSource( fModel->inboxIndex() );
                if ( inboxIndex.isValid() )
                    fImpl->folders->expand( inboxIndex );
                resizeToContentZero( fImpl->folders, EExpandMode::eNoAction );
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

void CFoldersView::reloadJunk()
{
    fModel->reloadJunk();
}

void CFoldersView::reloadTrash()
{
    fModel->reloadTrash();
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
    slotItemSelected();
}

QModelIndex CFoldersView::sourceIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() || ( idx.model() == fModel ) )
        return idx;
    return fFilterModel->mapToSource( idx );
}

QModelIndex CFoldersView::selectedIndex() const
{
    if ( !fImpl->folders->selectionModel() )
        return {};

    auto selectedIndexes = fImpl->folders->selectionModel()->selectedIndexes();
    if ( selectedIndexes.isEmpty() )
        return {};
    auto selectedIndex = selectedIndexes.first();
    if ( !selectedIndex.isValid() )
        return selectedIndex;
    return sourceIndex( selectedIndex );
}

void CFoldersView::slotSetRootFolder()
{
    auto idx = selectedIndex();
    if ( !idx.isValid() )
        return;
    auto folder = fModel->folderForIndex( idx );
    if ( !folder )
        return;
    COutlookAPI::instance()->setRootFolder( folder );
}

void CFoldersView::slotRunningStateChanged( bool running )
{
    fImpl->setRootFolderBtn->setEnabled( !running );
    fImpl->addFolder->setEnabled( !running );
    if ( !running )
        updateButtons( selectedIndex() );
}

void CFoldersView::slotItemSelected()
{
    auto index = selectedIndex();
    updateButtons( index );

    auto path = index.isValid() ? fModel->pathForIndex( index ) : QString();
    emit sigFolderSelected( path );
}

void CFoldersView::updateButtons( const QModelIndex &index )
{
    auto path = index.isValid() ? fModel->pathForIndex( index ) : QString();

    fImpl->setRootFolderBtn->setEnabled( index.isValid() );
    fImpl->addFolder->setEnabled( index.isValid() );
}

void CFoldersView::addFolder( const QString &folderName )
{
    auto newIndex = fFilterModel->mapFromSource( fModel->addFolder( selectedIndex(), folderName ) );
    selectAndScroll( newIndex );
}

void CFoldersView::slotAddFolder()
{
    auto newIndex = fFilterModel->mapFromSource( fModel->addFolder( selectedIndex(), this ) );
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
    auto idx = selectedIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->pathForIndex( idx );
}

QString CFoldersView::selectedFullPath() const
{
    auto idx = selectedIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->fullPathForIndex( idx );
}

COutlookObj< Outlook::MAPIFolder > CFoldersView::selectedFolder() const
{
    auto idx = selectedIndex();
    if ( !idx.isValid() )
        return {};
    return fModel->folderForIndex( idx );
}
