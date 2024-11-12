#include "EmailGroupingModel.h"
#include "OutlookHelpers.h"

#include "MSOUTL.h"

#include <algorithm>

CEmailGroupingModel::CEmailGroupingModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    setHorizontalHeaderLabels( QStringList() << "Domain" );
}

CEmailGroupingModel::~CEmailGroupingModel()
{
}

void CEmailGroupingModel::addEmailAddress( const QString &emailAddress )
{
    auto split = emailAddress.splitRef( '@', QString::SplitBehavior::SkipEmptyParts );
    if ( split.empty() )
        return;

    auto user = split.front();
    QStringRef domain;
    if ( split.count() == 2 )
    {
        domain = split.back();
    }
    else
        return;

    auto list = domain.split( '.', QString::SplitBehavior::SkipEmptyParts );
    std::reverse( std::begin( list ), std::end( list ) );
    list.push_back( user );
    findOrAddEmailAddressSection( list.front(), list.mid( 1 ), nullptr );
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
            parent->appendRow( item );
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
            appendRow( item );
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