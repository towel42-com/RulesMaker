#include "FoldersModel.h"
#include "OutlookAPI.h"

#include "MSOUTL.h"

#include <QTimer>
#include <QInputDialog>

#include <QMetaMethod>
#include <QDebug>

struct SCurrFolderInfo
{
    SCurrFolderInfo( const std::list< std::shared_ptr< Outlook::Folder > > &folders ) :
        fSubFolders( folders )
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

private:
    std::list< std::shared_ptr< Outlook::Folder > > fSubFolders;
    std::list< std::shared_ptr< Outlook::Folder > >::iterator fPos;
};

CFoldersModel::CFoldersModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
}

void CFoldersModel::reload()
{
    COutlookAPI::getInstance()->slotClearCanceled();

    auto folder = COutlookAPI::getInstance()->getInbox();
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

    auto sharedFolder = COutlookAPI::getInstance()->findMailFolder( folder );
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
    auto folder = COutlookAPI::getInstance()->getInbox();
    if ( !folder )
        return;

    QTimer::singleShot( 0, [ = ]() { addSubFolders( folder ); } );
}

void CFoldersModel::addSubFolders( const std::shared_ptr< Outlook::Folder > &rootFolder )
{
    fNumFolders = COutlookAPI::getInstance()->subFolderCount( rootFolder, true );
    fCurrFolderNum = 0;

    auto rootItem = new QStandardItem( COutlookAPI::getInstance()->folderName( rootFolder ) );
    appendRow( rootItem );
    fFolderMap[ rootItem ] = rootFolder;

    addSubFolders( rootItem, rootFolder );
}

void CFoldersModel::addSubFolders( QStandardItem *parent, const std::shared_ptr< Outlook::Folder > &parentFolder )
{
    auto curr = std::make_unique< SCurrFolderInfo >( COutlookAPI::getInstance()->getFolders( parentFolder, false ) );
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

    auto child = new QStandardItem( curr->folder()->Name() );
    fFolderMap[ child ] = curr->folder();
    parent->appendRow( child );
    emit sigSetStatus( ++fCurrFolderNum, fNumFolders );

    addSubFolders( child, curr->folder() );
    curr->incPos();
    if ( curr->atEnd() )
    {
        fFolders.erase( pos );

        parent->sortChildren( 0, Qt::SortOrder::AscendingOrder );
        if ( fFolders.empty() )
            emit sigFinishedLoading();
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
    auto parentFolder = COutlookAPI::getInstance()->getInbox();

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
