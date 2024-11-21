#include "FoldersModel.h"
#include "OutlookHelpers.h"

#include "MSOUTL.h"

#include <QTimer>
#include <QProgressDialog>
#include <QProgressBar>
#include <QInputDialog>

#include <QMetaMethod>
#include <QDebug>
CFoldersModel::CFoldersModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
}

void CFoldersModel::reload()
{
    auto folder = COutlookHelpers::getInstance()->getInbox( dynamic_cast< QWidget * >( parent() ) );
    if ( !folder )
        return;

    auto folders = new Outlook::Folders( folder->Folders() );

    //dumpMetaMethods( folder.get() );
    //dumpMetaMethods( folders );

    //connect( folders, SIGNAL( FolderAdd( Outlook::Folder * ) ), this, SLOT( slotAddFolder( Folder * ) ) );
    //connect( folders, SIGNAL( FolderChange( Folder * ) ), this, SLOT( slotFolderChanged( Folder * ) ) );
    //connect( folders, SIGNAL( FolderRemove() ), this, SLOT( slotReload() ) );

    QTimer::singleShot( 0, [ = ]() { slotReload(); } );
}

CFoldersModel::~CFoldersModel()
{
}

void CFoldersModel::slotAddFolder( Outlook::Folder *folder )
{
    if ( !folder )
        return;

    auto sharedFolder = std::shared_ptr< Outlook::Folder >( folder );
    auto child = new QStandardItem( folder->Name() );
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
    auto folder = COutlookHelpers::getInstance()->getInbox( dynamic_cast< QWidget * >( parent() ) );
    if ( !folder )
        return;

    QTimer::singleShot( 0, [ = ]() { addSubFolders( folder ); } );
}

void CFoldersModel::addSubFolders( std::shared_ptr< Outlook::Folder > rootFolder )
{
    auto subFoldersSize = static_cast< int >( COutlookHelpers::getInstance()->getFolders( rootFolder, false ).size() );

    QProgressDialog dlg( dynamic_cast< QWidget * >( parent() ) );
    auto bar = new QProgressBar;
    bar->setFormat( "(%v of %m - %p%)" );
    dlg.setBar( bar );
    dlg.setMinimum( 0 );
    dlg.setMaximum( subFoldersSize );
    dlg.setLabelText( "Loading Folders" );
    dlg.setMinimumDuration( 0 );
    dlg.setWindowModality( Qt::WindowModal );

    auto rootItem = new QStandardItem( rootFolder->Name() );
    appendRow( rootItem );
    fFolderMap[ rootItem ] = rootFolder;

    if ( !addSubFolders( rootItem, rootFolder, &dlg ) )
    {
        clear();
        return;
    }
    rootItem->sortChildren( 0, Qt::SortOrder::AscendingOrder );
    emit sigFinishedLoading();
}

bool CFoldersModel::addSubFolders( QStandardItem *parent, std::shared_ptr< Outlook::Folder > parentFolder, QProgressDialog *progress )
{
    auto subFolders = COutlookHelpers::getInstance()->getFolders( parentFolder, false );
    for ( auto &&ii : subFolders )
    {
        if ( progress )
        {
            progress->setValue( progress->value() + 1 );
            if ( progress->wasCanceled() )
            {
                return false;
            }
        }
        auto folder = ii;
        auto child = new QStandardItem( ii->Name() );
        fFolderMap[ child ] = folder;
        parent->appendRow( child );
        addSubFolders( child, folder, nullptr );
    }
    parent->sortChildren( 0, Qt::SortOrder::AscendingOrder );
    return true;
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
    setHorizontalHeaderLabels( QStringList() << "Folder Name" );
    fFolderMap.clear();
}

void CFoldersModel::addFolder( const QModelIndex &idx, QWidget *parent )
{
    auto parentFolder = COutlookHelpers::getInstance()->getInbox( parent );

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
