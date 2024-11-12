#include "ContactsModel.h"
#include "OutlookHelpers.h"

#include "msoutl.h"

CContactsModel::CContactsModel( QObject *parent ) :
    QAbstractListModel( parent )
{
    auto folder = COutlookHelpers::getInstance()->getContacts( dynamic_cast< QWidget * >( parent ) );
    if ( !folder )
        return;

    fItems = std::make_unique< Outlook::Items >( folder->Items() );
    if ( fItems )
        fCountCache = fItems->Count();

    connect( fItems.get(), SIGNAL( ItemAdd( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( fItems.get(), SIGNAL( ItemChange( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( fItems.get(), SIGNAL( ItemRemove() ), parent, SLOT( updateOutlook() ) );
}

CContactsModel::~CContactsModel()
{
}

int CContactsModel::rowCount( const QModelIndex & ) const
{
    if ( fItems && fCountCache.has_value() )
        return fCountCache.value();
    return fItems ? fItems->Count() : 0;
}

int CContactsModel::columnCount( const QModelIndex & /*parent*/ ) const
{
    return 4;
}

QVariant CContactsModel::headerData( int section, Qt::Orientation /*orientation*/, int role ) const
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

QVariant CContactsModel::data( const QModelIndex &index, int role ) const
{
    if ( !index.isValid() || role != Qt::DisplayRole )
        return QVariant();

    QStringList data;
    if ( fCache.contains( index.row() ) )
    {
        data = fCache.value( index.row() );
    }
    else if ( fItems )
    {
        Outlook::ContactItem contact( fItems->Item( index.row() + 1 ) );
        data << contact.FirstName() << contact.LastName() << contact.HomeAddress() << contact.Email1Address();
        fCache.insert( index.row(), data );
    }

    if ( index.column() < data.count() )
        return data.at( index.column() );

    return QVariant();
}

void CContactsModel::changeItem( const QModelIndex &index, const QString &firstName, const QString &lastName, const QString &address, const QString &email )
{
    if ( !fItems )
        return;

    Outlook::ContactItem item( fItems->Item( index.row() + 1 ) );

    item.SetFirstName( firstName );
    item.SetLastName( lastName );
    item.SetHomeAddress( address );
    item.SetEmail1Address( email );

    item.Save();

    fCache.take( index.row() );
}

void CContactsModel::addItem( const QString &firstName, const QString &lastName, const QString &address, const QString &email )
{
    Outlook::ContactItem item( COutlookHelpers::getInstance()->outlook()->CreateItem( Outlook::OlItemType::olContactItem ) );
    if ( !item.isNull() )
    {
        item.SetFirstName( firstName );
        item.SetLastName( lastName );
        item.SetHomeAddress( address );
        item.SetEmail1Address( email );

        item.Save();
    }
}

void CContactsModel::update()
{
    beginResetModel();
    fCache.clear();
    endResetModel();
}

