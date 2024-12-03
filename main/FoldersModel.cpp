#include "FoldersModel.h"
#include "OutlookAPI.h"

#include "MSOUTL.h"

#include <QTimer>
#include <QInputDialog>

#include <QMetaMethod>
#include <QDebug>

struct SCurrFolderInfo
{
    SCurrFolderInfo( const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool isRoot ) :
        fSubFolders( folders ),
        fIsRoot( isRoot )
    {
        fPos = fSubFolders.begin();
    }
    std::shared_ptr< Outlook::Folder > folder() const
    {
        if ( fPos == fSubFolders.end() )
            return {};
        return ( *fPos );
    }
    void incPos()
    {
        if ( fPos == fSubFolders.end() )
            return;
        fPos++;
    }
    bool empty() const { return fSubFolders.empty(); }
    bool atEnd() const { return fPos == fSubFolders.end(); }

    bool isRoot() const { return fIsRoot; }

private:
    std::list< std::shared_ptr< Outlook::Folder > > fSubFolders;
    std::list< std::shared_ptr< Outlook::Folder > >::iterator fPos;
    bool fIsRoot{ false };
};

CFoldersModel::CFoldersModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
}

void CFoldersModel::reload()
{
    COutlookAPI::getInstance()->slotClearCanceled();

    auto folder = COutlookAPI::getInstance()->getInbox( dynamic_cast< QWidget * >( parent() ) );
    if ( !folder )
        return;

    QTimer::singleShot( 0, [ = ]() { slotReload(); } );
}

CFoldersModel::~CFoldersModel()
{
}

void CFoldersModel::slotAddFolder( Outlook::Folder *folder )
{
    if ( !folder )
        return;

    auto sharedFolder = COutlookAPI::getInstance()->getFolder( folder );
    auto child = new QStandardItem( COutlookAPI::getInstance()->folderName( folder ) );
    fFolderMap[ child ] = sharedFolder;

    QStandardItem *parentItem = nullptr;
    auto parent = folder->Parent();
    if ( parent )
    {
        auto parentObj = new Outlook::Folder( parent );
        if ( parentObj->Class() == Outlook::OlObjectClass::olFolder )
        {
            auto parentPath = parentObj->FullFolderPath();
            for ( auto &&ii : fFolderMap )
            {
                if ( ii.second->FullFolderPath() == parentPath )
                {
                    parentItem = ii.first;
                    break;
                }
            }
        }
        delete parentObj;
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
}

void CFoldersModel::slotFolderChanged( Outlook::Folder * /*folder*/ )
{
}

void CFoldersModel::slotReload()
{
    clear();
    auto folder = COutlookAPI::getInstance()->getInbox( dynamic_cast< QWidget * >( parent() ) );
    if ( !folder )
        return;

    QTimer::singleShot( 0, [ = ]() { addSubFolders( folder ); } );
}

void CFoldersModel::addSubFolders( const std::shared_ptr< Outlook::Folder > &rootFolder )
{
    auto subFoldersSize = COutlookAPI::getInstance()->subFolderCount( rootFolder );
    auto rootItem = new QStandardItem( COutlookAPI::getInstance()->folderName( rootFolder ) );
    appendRow( rootItem );
    fFolderMap[ rootItem ] = rootFolder;

    emit sigSetStatus( 0, subFoldersSize );
    if ( COutlookAPI::getInstance()->canceled() )
    {
        clear();
        emit sigFinishedLoading();
        return;
    }

    addSubFolders( rootItem, rootFolder, true );
}

void CFoldersModel::addSubFolders( QStandardItem *parent, const std::shared_ptr< Outlook::Folder > &parentFolder, bool root )
{
    auto curr = std::make_unique< SCurrFolderInfo >( COutlookAPI::getInstance()->getFolders( parentFolder, false ), root );
    if ( curr->atEnd() )
        return;

    fFolders[ parent ] = std::move( curr );

    QTimer::singleShot( 0, [ = ]() { slotAddNextFolder( parent ); } );
}

void CFoldersModel::slotAddNextFolder( QStandardItem *parent )
{
    if ( COutlookAPI::getInstance()->canceled() )
    {
        clear();
        emit sigFinishedLoading();
        return;
    }

    auto pos = fFolders.find( parent );
    if ( pos == fFolders.end() )
        return;

    auto &&curr = ( *pos ).second;
    if ( curr->isRoot() )
        emit sigIncStatusValue();

    auto child = new QStandardItem( curr->folder()->Name() );
    fFolderMap[ child ] = curr->folder();
    parent->appendRow( child );
    addSubFolders( child, curr->folder(), false );
    curr->incPos();
    if ( !curr->atEnd() )
        QTimer::singleShot( 0, [ = ]() { slotAddNextFolder( parent ); } );
    else
    {
        if ( COutlookAPI::getInstance()->canceled() )
        {
            clear();
            emit sigFinishedLoading();
            return;
        }

        fFolders.erase( pos );

        parent->sortChildren( 0, Qt::SortOrder::AscendingOrder );
        if ( fFolders.empty() )
            emit sigFinishedLoading();
    }
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
    return folder->FullFolderPath();
}

void CFoldersModel::clear()
{
    QStandardItemModel::clear();
    fFolders.clear();
    setHorizontalHeaderLabels( QStringList() << "Folder Name" );
    fFolderMap.clear();
}

void CFoldersModel::addFolder( const QModelIndex &idx, QWidget *parent )
{
    auto parentFolder = COutlookAPI::getInstance()->getInbox( parent );

    auto folderName = QInputDialog::getText( parent, "New Folder Name", "Folder Name" );
    if ( folderName.isEmpty() )
        return;

    auto parentItem = itemFromIndex( idx );
    if ( parentItem )
    {
        auto pos = fFolderMap.find( parentItem );
        if ( pos != fFolderMap.end() )
            parentFolder = ( *pos ).second;
    }

    if ( !parentFolder )
        return;

    auto newFolder = parentFolder->Folders()->Add( folderName );
    slotAddFolder( reinterpret_cast< Outlook::Folder * >( newFolder ) );
}
