#include "FoldersModel.h"
#include "OutlookHelpers.h"

#include "MSOUTL.h"

#include <QTimer>
#include <QProgressDialog>
#include <QProgressBar>

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
    QTimer::singleShot( 0, [ = ]() { addSubFolders( folder ); } );
}

CFoldersModel::~CFoldersModel()
{
}

void CFoldersModel::addSubFolders( std::shared_ptr< Outlook::MAPIFolder > rootFolder )
{
    auto rootItem = new QStandardItem( rootFolder->Name() );
    appendRow( rootItem );

    auto subFolders = COutlookHelpers::getInstance()->getFolders( rootFolder, false );

    QProgressDialog dlg( dynamic_cast< QWidget * >( parent() ) );
    auto bar = new QProgressBar;
    bar->setFormat( "(%v of %m - %p%)" );
    dlg.setBar( bar );
    dlg.setMinimum( 0 );
    dlg.setMaximum( static_cast< int >( subFolders.size() ) );
    dlg.setLabelText( "Loading Folders" );
    dlg.setMinimumDuration( 0 );
    dlg.setWindowModality( Qt::WindowModal );

    for ( auto &&ii : subFolders )
    {
        dlg.setValue( dlg.value() + 1 );
        if ( dlg.wasCanceled() )
        {
            clear();
            return;
        }
        auto child = new QStandardItem( ii->Name() );
        rootItem->appendRow( child );
        addSubFolders( child, ii );
    }
    rootItem->sortChildren( 0, Qt::SortOrder::AscendingOrder );
    emit sigFinishedLoading();
}

void CFoldersModel::addSubFolders( QStandardItem *parent, std::shared_ptr< Outlook::MAPIFolder > parentFolder )
{
    auto subFolders = COutlookHelpers::getInstance()->getFolders( parentFolder, false );
    for ( auto &&ii : subFolders )
    {
        auto child = new QStandardItem( ii->Name() );
        parent->appendRow( child );
        addSubFolders( child, ii );
    }
    parent->sortChildren( 0, Qt::SortOrder::AscendingOrder );
}

QString CFoldersModel::fullPath( const QModelIndex &index ) const
{
    auto item = this->itemFromIndex( index );
    return fullPath( item );
}

QString CFoldersModel::fullPath( QStandardItem *item ) const
{
    if ( !item )
        return {};
    QString retVal = item->text();
    auto parent = item->parent();
    if ( parent )
    {
        auto parentPath = fullPath( parent );
        if ( !parentPath.isEmpty() )
            retVal = parentPath + R"(\)" + retVal;
    }
    else
        retVal = R"(\\)" + retVal;

    return retVal;
}

void CFoldersModel::clear()
{
    QStandardItemModel::clear();
    setHorizontalHeaderLabels( QStringList() << "Folder Name" );
}
