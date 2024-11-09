#include "RulesModel.h"
#include "OutlookHelpers.h"

#include "msoutl.h"

CRulesModel::CRulesModel( QObject *parent ) :
    QAbstractListModel( parent )
{
    auto folder = COutlookHelpers::getInstance()->selectContactFolder( dynamic_cast< QWidget * >( parent ) );
    if ( !folder )
        return;

    fItems = std::make_unique< Outlook::Items >( folder->Items() );
    connect( fItems.get(), SIGNAL( ItemAdd( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( fItems.get(), SIGNAL( ItemChange( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( fItems.get(), SIGNAL( ItemRemove() ), parent, SLOT( updateOutlook() ) );
}

CRulesModel::~CRulesModel()
{
}

int CRulesModel::rowCount( const QModelIndex & ) const
{
    return fItems ? fItems->Count() : 0;
}

int CRulesModel::columnCount( const QModelIndex & /*parent*/ ) const
{
    return 4;
}

QVariant CRulesModel::headerData( int section, Qt::Orientation /*orientation*/, int role ) const
{
    if ( role != Qt::DisplayRole )
        return QVariant();

    switch ( section )
    {
        case 0:
            return tr( "First Name" );
        case 1:
            return tr( "Last Name" );
        case 2:
            return tr( "Address" );
        case 3:
            return tr( "Email" );
        default:
            break;
    }

    return QVariant();
}

QVariant CRulesModel::data( const QModelIndex &index, int role ) const
{
    if ( !index.isValid() || role != Qt::DisplayRole )
        return QVariant();

    QStringList data;
    if ( fCache.contains( index ) )
    {
        data = fCache.value( index );
    }
    else if ( fItems )
    {
        Outlook::Rule rule( fItems->Item( index.row() + 1 ) );
        data << rule.Name() << QString() << QString();
        fCache.insert( index, data );
    }

    if ( index.column() < data.count() )
        return data.at( index.column() );

    return QVariant();
}

void CRulesModel::changeItem( const QModelIndex &/*index*/, const QString &/*folderName*/ )
{
    if ( !fItems )
        return;

    //Outlook::Folder item( fItems->Item( index.row() + 1 ) );

    //item.SetName( folderName );
    ////item.Save();

    //fCache.take( index );
}

void CRulesModel::addItem( const QString & /*folderName*/ )
{
    //Outlook::Folder item( COutlookHelpers::getInstance()->outlook()->CreateItem( Outlook::OlItemType::olContactItem ) );
    //if ( !item.isNull() )
    //{
    //    item.SetName( folderName );
    //    item.Save();
    //}
}

void CRulesModel::update()
{
    beginResetModel();
    fCache.clear();
    endResetModel();
}
