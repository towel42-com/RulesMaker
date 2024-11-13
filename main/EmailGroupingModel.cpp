#include "EmailGroupingModel.h"
#include "OutlookHelpers.h"

#include "MSOUTL.h"

#include <algorithm>

CEmailGroupingModel::CEmailGroupingModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
}

CEmailGroupingModel::~CEmailGroupingModel()
{
}

void CEmailGroupingModel::addEmailAddress( std::shared_ptr< Outlook::MailItem > mailItem, const QString &emailAddress )
{
    auto pos = fCache.find( emailAddress );
    if ( pos != fCache.end() )
        return;

    auto split = emailAddress.splitRef( '@', QString::SplitBehavior::SkipEmptyParts );
    if ( split.empty() )
        return;

    auto user = split.front();
    QStringRef domain;
    if ( split.count() != 2 )
        return;

    domain = split.back();
    pos = fDomainCache.find( domain.toString() );
    if ( pos != fDomainCache.end() )
    {
        auto retVal = findOrAddEmailAddressSection( user, {}, ( *pos ).second );
        fCache[ emailAddress ] = retVal;
        fEmailCache[ retVal ] = mailItem;
        return;
    }

    auto list = domain.split( '.', QString::SplitBehavior::SkipEmptyParts );
    std::reverse( std::begin( list ), std::end( list ) );
    list.push_back( user );
    auto retVal = findOrAddEmailAddressSection( list.front(), list.mid( 1 ), nullptr );
    if ( retVal )
    {
        fCache[ emailAddress ] = retVal;
        fEmailCache[ retVal ] = mailItem;
        auto parent = dynamic_cast< CEmailAddressSection * >( retVal->parent() );
        if ( parent )
            fDomainCache[ domain.toString() ] = parent;
    }
}

void CEmailGroupingModel::clear()
{
    QStandardItemModel::clear();
    setHorizontalHeaderLabels(
        QStringList() << "Domain"
                      << "User" );
    beginResetModel();
    fRootItems.clear();
    fCache.clear();
    fDomainCache.clear();
    endResetModel();
}

std::shared_ptr< Outlook::MailItem > CEmailGroupingModel::emailItemFromIndex( const QModelIndex &idx )
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    if ( !item )
        return {};
    auto pos = fEmailCache.find( item );
    if ( pos == fEmailCache.end() )
    {
        if ( idx.column() == 0 )
        {
            auto otherIdx = this->index( idx.row(), 1, idx.parent() );
            return emailItemFromIndex( otherIdx );
        }
        return {};
    }
    return ( *pos ).second;
}

QList< QStandardItem * > makeRow( QStandardItem *item, bool inBack )
{
    QList< QStandardItem * > row;
    if ( inBack )
        row << new CEmailAddressSection << item;
    else
        row << item << new CEmailAddressSection;
    return row;
}

CEmailAddressSection *CEmailGroupingModel::findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent )
{
    CEmailAddressSection *retVal{ nullptr };
    auto key = curr.toString().toLower();
    if ( parent )
    {
        auto pos = parent->fChildItems.find( key );
        if ( pos == parent->fChildItems.end() )
        {
            auto item = new CEmailAddressSection( key );

            parent->appendRow( makeRow( item, remaining.empty() ) );
            parent->fChildItems[ key ] = item;
            retVal = item;
        }
        else
            retVal = ( *pos ).second;
    }
    else
    {
        auto pos = fRootItems.find( key );
        if ( pos == fRootItems.end() )
        {
            auto item = new CEmailAddressSection( key );
            appendRow( makeRow( item, remaining.empty() ) );
            fRootItems[ key ] = item;
            retVal = item;
        }
        else
            retVal = ( *pos ).second;
    }

    if ( !remaining.empty() )
        return findOrAddEmailAddressSection( remaining.front(), remaining.mid( 1 ), retVal );
    return retVal;
}