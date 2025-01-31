#include "ListFilterModel.h"

CListFilterModel::CListFilterModel( QObject *parent ) :
    QSortFilterProxyModel( parent )
{
    setDynamicSortFilter( false );
    setFilterCaseSensitivity( Qt::CaseInsensitive );
    setRecursiveFilteringEnabled( true );
}

CListFilterModel::~CListFilterModel()
{
}

bool CListFilterModel::filterAcceptsColumn( int source_column, const QModelIndex &source_parent ) const
{
    return QSortFilterProxyModel::filterAcceptsColumn( source_column, source_parent );
}

bool CListFilterModel::filterAcceptsRow( int sourceRow, const QModelIndex &sourceParent ) const
{
    if ( onlyFilterParent() && sourceParent.isValid() )
        return true;

    auto retVal = QSortFilterProxyModel::filterAcceptsRow( sourceRow, sourceParent );
    if ( retVal && fShowRowFunc )
        retVal = fShowRowFunc( sourceRow, sourceParent );

    return retVal;
}

bool CListFilterModel::lessThan( const QModelIndex &source_left, const QModelIndex &source_right ) const
{
    if ( fLessThanOp )
        return fLessThanOp( source_left, source_right );
    return QSortFilterProxyModel::lessThan( source_left, source_right );
}

void CListFilterModel::slotSetFilter( const QString &filter )
{
    if ( filter.isEmpty() )
        setFilterWildcard( filter );
    else
        setFilterWildcard( "*" + filter + "*" );
}
