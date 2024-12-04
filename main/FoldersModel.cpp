#include "FoldersModel.h"
#include "OutlookAPI.h"

#include "MSOUTL.h"

#include <QTimer>
#include <QInputDialog>
#include <QSortFilterProxyModel>

#include <QMetaMethod>
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
    fNumFolders = COutlookAPI::getInstance()->recursiveSubFolderCount( rootFolder.get() );
    fCurrFolderNum = 0;

    auto rootItem = new QStandardItem( COutlookAPI::getInstance()->folderName( rootFolder ) );
    appendRow( rootItem );
    fFolderMap[ rootItem ] = rootFolder;

    addSubFolders( rootItem, rootFolder );
}

void CFoldersModel::addSubFolders( QStandardItem *parent, const std::shared_ptr< Outlook::Folder > &parentFolder )
{
    auto folders = COutlookAPI::getInstance()->getFolders( parentFolder, false );
    auto curr = std::make_unique< SCurrFolderInfo >( TFolderVector( { folders.begin(), folders.end() } ) );
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

    auto parentName = parent ? parent->text() : QString();

    auto folderName = curr->folder()->Name();
    auto child = new QStandardItem( /*QString::number( curr->pos() )  + " - " + */ folderName );
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
