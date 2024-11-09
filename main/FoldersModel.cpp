#include "FoldersModel.h"
#include "OutlookHelpers.h"

#include "msoutl.h"

CFoldersModel::CFoldersModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    auto folder = COutlookHelpers::getInstance()->selectInboxFolder( dynamic_cast< QWidget * >( parent ) );
    if ( !folder )
        return;

    setHorizontalHeaderLabels( QStringList() << "Folder Name" );
    auto rootItem = new QStandardItem( folder->Name() );
    appendRow( rootItem );
    addSubFolders( rootItem, folder );

    connect( folder.get(), SIGNAL( ItemAdd( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( folder.get(), SIGNAL( ItemChange( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( folder.get(), SIGNAL( ItemRemove() ), parent, SLOT( updateOutlook() ) );
}

CFoldersModel::~CFoldersModel()
{
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

void CFoldersModel::changeItem( const QModelIndex & /*index*/, const QString & /*folderName*/ )
{
    //if ( !fFolders )
    //    return;

    //Outlook::Folder item( fFolders->Item( index.row() + 1 ) );

    //item.SetName( folderName );
    ////item.Save();

    //fCache.take( index );
}

void CFoldersModel::addItem( const QString & /*folderName*/ )
{
    //Outlook::Folder item( COutlookHelpers::getInstance()->outlook()->CreateItem( Outlook::OlItemType::olContactItem ) );
    //if ( !item.isNull() )
    //{
    //    item.SetName( folderName );
    //    item.Save();
    //}
}

void CFoldersModel::update()
{
    beginResetModel();
    //fCache.clear();
    endResetModel();
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
