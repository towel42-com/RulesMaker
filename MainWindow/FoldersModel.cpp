#include "FoldersModel.h"
#include "OutlookAPI/OutlookAPI.h"

#include <QTimer>
#include <QInputDialog>
#include <QSortFilterProxyModel>

#include <QDebug>
#include <vector>

using TFolderVector = std::vector< std::shared_ptr< Outlook::Folder > >;

struct SCurrFolderInfo
{
    SCurrFolderInfo( const TFolderVector &folders ) :
        fSubFolders( folders )
    {
        fPos = 0;
    }
    std::shared_ptr< Outlook::Folder > folder() const
    {
        if ( atEnd() )
            return {};
        return fSubFolders[ fPos ];
    }
    void incPos()
    {
        if ( atEnd() )
            return;

        fPos++;
    }
    bool empty() const { return fSubFolders.empty(); }
    bool atEnd() const { return fPos >= fSubFolders.size(); }
    int pos() const { return static_cast< int >( fPos ); }

private:
    TFolderVector fSubFolders;
    std::size_t fPos{ 0 };
};

CFoldersModel::CFoldersModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
}

void CFoldersModel::reload()
{
    COutlookAPI::instance()->slotClearCanceled();

    QTimer::singleShot( 0, [ = ]() { slotReload(); } );
}

CFoldersModel::~CFoldersModel()
{
}

void CFoldersModel::slotReload()
{
    clear();
    auto inbox = COutlookAPI::instance()->getInbox();
    auto junk = COutlookAPI::instance()->getJunkFolder();
    QTimer::singleShot( 0, [ = ]() { loadRootFolders( { inbox, junk } ); } );
}

void CFoldersModel::loadRootFolders( const std::list< std::shared_ptr< Outlook::Folder > > &rootFolders )
{
    fCurrFolderNum = 0;
    fNumFolders = 0;
    for ( auto &&ii : rootFolders )
    {
        fNumFolders += COutlookAPI::instance()->recursiveSubFolderCount( ii.get() );

        auto rootItem = addFolder( ii, nullptr );
        addSubFolders( rootItem, ii );
    }
}

void CFoldersModel::addSubFolders( QStandardItem *parent, const std::shared_ptr< Outlook::Folder > &parentFolder )
{
    auto folders = COutlookAPI::instance()->getFolders( parentFolder, false );
    auto curr = std::make_unique< SCurrFolderInfo >( TFolderVector( { folders.begin(), folders.end() } ) );
    if ( curr->atEnd() )
        return;

    fFolders[ parent ] = std::move( curr );

    QTimer::singleShot( 0, [ = ]() { slotAddNextFolder( parent ); } );
}

void CFoldersModel::slotAddNextFolder( QStandardItem *parent )
{
    if ( COutlookAPI::instance()->canceled() )
    {
        clear();
        emit sigFinishedLoading();
        return;
    }

    auto pos = fFolders.find( parent );
    if ( pos == fFolders.end() )
        return;

    auto &&curr = ( *pos ).second;

    auto parentName = parent ? parent->text() : QString();

    auto folderName = COutlookAPI::instance()->nameForFolder( curr->folder() );
    auto child = new QStandardItem( folderName );
    fFolderMap[ child ] = curr->folder();
    parent->insertRow( curr->pos(), child );
    emit sigSetStatus( ++fCurrFolderNum, fNumFolders );

    addSubFolders( child, curr->folder() );
    curr->incPos();
    if ( curr->atEnd() )
    {
        fFolders.erase( pos );
        if ( fFolders.empty() )
            emit sigFinishedLoading();
        else
            emit sigFinishedLoadingChildren( parent );
    }
    else
        QTimer::singleShot( 0, [ = ]() { slotAddNextFolder( parent ); } );
}

std::shared_ptr< Outlook::Folder > CFoldersModel::folderForItem( const QModelIndex &index ) const
{
    auto item = this->itemFromIndex( index );
    return folderForItem( item );
}

std::shared_ptr< Outlook::Folder > CFoldersModel::folderForItem( QStandardItem *item ) const
{
    auto pos = fFolderMap.find( item );
    if ( pos == fFolderMap.end() )
        return {};
    return ( *pos ).second;
}

QString CFoldersModel::pathForItem( const QModelIndex &index ) const
{
    auto item = this->itemFromIndex( index );
    return pathForItem( item );
}

QString CFoldersModel::pathForItem( QStandardItem *item ) const
{
    if ( !item )
        return {};
    QString retVal = item->text();
    auto parent = item->parent();
    if ( parent )
    {
        auto parentPath = pathForItem( parent );
        if ( !parentPath.isEmpty() )
            retVal = parentPath + R"(\)" + retVal;
    }
    else
        retVal = R"(\\)" + retVal;

    return retVal;
}

QString CFoldersModel::fullPathForItem( const QModelIndex &index ) const
{
    auto item = this->itemFromIndex( index );
    return fullPathForItem( item );
}

QString CFoldersModel::fullPathForItem( QStandardItem *item ) const
{
    auto folder = folderForItem( item );
    if ( !folder )
        return {};
    return COutlookAPI::instance()->rawPathForFolder( folder );
}

void CFoldersModel::clear()
{
    QStandardItemModel::clear();
    fFolders.clear();
    setHorizontalHeaderLabels( QStringList() << "Folder Name" );
    fFolderMap.clear();
}

QModelIndex CFoldersModel::addFolder( const QModelIndex &parentIndex, QWidget *parent )
{
    auto folderName = QInputDialog::getText( parent, "New Folder Name", "Folder Name" );
    if ( folderName.isEmpty() )
        return {};

    return addFolder( parentIndex, folderName );
}

QModelIndex CFoldersModel::addFolder( const QModelIndex &parentIndex, const QString &folderName )
{
    auto parentFolder = COutlookAPI::instance()->getInbox();

    auto parentItem = itemFromIndex( parentIndex );
    if ( parentItem )
    {
        auto pos = fFolderMap.find( parentItem );
        if ( pos != fFolderMap.end() )
            parentFolder = ( *pos ).second;
    }

    if ( !parentFolder )
        return {};

    auto newFolder = COutlookAPI::instance()->addFolder( parentFolder, folderName );

    auto retVal = addFolder( newFolder, parentItem );
    if ( !retVal )
        return {};
    return indexFromItem( retVal );
}

QStandardItem *CFoldersModel::addFolder( const std::shared_ptr< Outlook::Folder > &folder, QStandardItem *parentItem )
{
    if ( !folder )
        return nullptr;

    auto api = COutlookAPI::instance();

    auto folderName = api->folderDisplayName( folder );
    auto child = new QStandardItem( folderName );
    fFolderMap[ child ] = folder;

    if ( !parentItem )
    {
        auto parentFolder = api->parentFolder( folder );
        if ( !parentFolder )
        {
            auto parentPath = api->rawPathForFolder( parentFolder );
            for ( auto &&ii : fFolderMap )
            {
                if ( api->rawPathForFolder( ii.second ) == parentPath )
                {
                    parentItem = ii.first;
                    break;
                }
            }
        }
    }
    
    if ( !parentItem )
    {
        appendRow( child );
        sort( 0, Qt::SortOrder::AscendingOrder );
    }
    else
    {
        parentItem->appendRow( child );
        parentItem->sortChildren( 0, Qt::SortOrder::AscendingOrder );
    }
    return child;
}
