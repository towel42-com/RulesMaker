#include "GroupedEmailModel.h"
#include "OutlookAPI.h"

#include "MSOUTL.h"

#include <QTimer>
#include <algorithm>

#ifdef _DEBUG
// #define LIMIT_EMAIL_READ
#endif
CGroupedEmailModel::CGroupedEmailModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
    connect( COutlookAPI::getInstance().get(), &COutlookAPI::sigOptionChanged, this, &CGroupedEmailModel::reload );
}

void CGroupedEmailModel::clear()
{
    QStandardItemModel::clear();
    setHorizontalHeaderLabels(
        QStringList() << "Domain"
                      << "User" );
    beginResetModel();
    fItems.reset();
    fRootItems.clear();
    fCache.clear();
    fDomainCache.clear();
    fEmailCache.clear();
    fCountCache.reset();
    fCurrPos = 1;
    endResetModel();
}

void CGroupedEmailModel::reload()
{
    COutlookAPI::getInstance()->slotClearCanceled();

    beginResetModel();
    clear();
    auto folder = COutlookAPI::getInstance()->rootProcessFolder();
    if ( !folder )
        return;
    auto fn = folder->FullFolderPath();

    auto items = folder->Items();
    if ( items )
    {
        auto limitToUnread = COutlookAPI::getInstance()->onlyProcessUnread();
        if ( limitToUnread && ( items->Count() < 200 ) )
            limitToUnread = !COutlookAPI::getInstance()->processAllEmailWhenLessThan200Emails();

        if ( limitToUnread )
        {
            auto subItems = items->Restrict( "[UNREAD]=TRUE" );
            if ( subItems )
                fItems = COutlookAPI::getInstance()->getItems( subItems );
        }
        if ( !fItems )
            fItems = COutlookAPI::getInstance()->getItems( items );
        if ( fItems )
            fCountCache = fItems->Count();
    }

    endResetModel();
    fCurrPos = 1;
    QTimer::singleShot( 0, [ = ]() { groupNextMailItemBySender(); } );
}

void CGroupedEmailModel::groupNextMailItemBySender()
{
    if ( !fItems || !fCountCache.has_value() )
        return;

    if ( fCurrPos > fCountCache.value() )
        return;

    if ( COutlookAPI::getInstance()->canceled() )
    {
        clear();
        emit sigFinishedGrouping();
        return;
    }

    auto item = fItems->Item( fCurrPos );
    if ( !item )
        return;

    if ( COutlookAPI::getObjectClass( item ) == Outlook::OlObjectClass::olMail )
    {
        auto mail = COutlookAPI::getInstance()->getMailItem( item );
        addEmailAddresses( mail, COutlookAPI::getSenderEmailAddresses( mail.get() ) );
    }

    emit sigSetStatus( fCurrPos, fCountCache.value() );

#ifdef LIMIT_EMAIL_READ
    if ( fCurrPos >= 100 )
        return;
#endif
    if ( fCurrPos == fCountCache.value() )
    {
        sortAll( nullptr );
        emit sigFinishedGrouping();
    }
    else
    {
        fCurrPos++;
        QTimer::singleShot( 0, [ = ]() { groupNextMailItemBySender(); } );
    }
}

CGroupedEmailModel::~CGroupedEmailModel()
{
}

void CGroupedEmailModel::sortAll( QStandardItem *root )
{
    if ( root )
        root->sortChildren( 0 );
    else
        sort( 0 );

    int count = root ? root->rowCount() : rowCount();
    for ( int ii = 0; ii < count; ++ii )
    {
        auto child = root ? root->child( ii ) : item( ii );
        sortAll( child );
    }
}

void CGroupedEmailModel::addEmailAddresses( std::shared_ptr< Outlook::MailItem > mailItem, const QStringList &emailAddresses )
{
    if ( !mailItem )
        return;

    for ( auto &&emailAddress : emailAddresses )
    {
        if ( emailAddress.indexOf( "bestdvibe" ) != -1 )
            int xyz = 0;
        addEmailAddress( mailItem, emailAddress );
    }
}

void CGroupedEmailModel::addEmailAddress( std::shared_ptr< Outlook::MailItem > mailItem, const QString &emailAddress )
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

int maxDepth( QStandardItem *item )
{
    int base = 0;
    if ( item->hasChildren() )
        base += 1;

    int retVal = base;
    for ( int ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii );
        auto curr = base + maxDepth( child );
        retVal = std::max( retVal, curr );
    }
    return retVal;
}

bool hasSingleLevelChild( QStandardItem *item )
{
    if ( !item->hasChildren() )
        return false;

    for ( int ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii );
        if ( !child->hasChildren() )
            return true;
    }
    return false;
}

QStringList CGroupedEmailModel::rulesForIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    return rulesForItem( item );
}

QStringList CGroupedEmailModel::rulesForItem( QStandardItem *item ) const
{
    QStringList retVal;

    auto rule = ruleForItem( item );
    if ( ( item->column() == 0 ) && ( rule.indexOf( '@' ) == -1 ) )
        rule = "@" + rule;
    retVal.push_back( rule );
    for ( auto &&ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii, 0 );
        if ( !child || child->text().isEmpty() )
            continue;
        //retVal.push_back( "@" + ruleForItem( child ) );
        auto curr = rulesForItem( child );
        retVal << curr;
    }
    return retVal;
}

QString CGroupedEmailModel::ruleForItem( QStandardItem *item ) const
{
    if ( !item )
        return {};
    QStringList path;
    QString separator = ".";
    auto colZeroItem = itemFromIndex( index( item->index().row(), 0, item->index().parent() ) );
    if ( !colZeroItem || !colZeroItem->hasChildren() )
        separator = "@";

    auto itemText = item->text();
    if ( itemText.isEmpty() )
    {
        int column = ( item->index().column() == 0 ) ? 1 : 0;
        itemText = itemFromIndex( index( item->index().row(), column, item->index().parent() ) )->text();
    }
    path.push_back( itemText );

    auto parent = item->parent();
    if ( parent )
        path.push_back( ruleForItem( parent ) );

    return path.join( separator );
}

std::shared_ptr< Outlook::MailItem > CGroupedEmailModel::emailItemFromIndex( const QModelIndex &idx ) const
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

CEmailAddressSection *CGroupedEmailModel::findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent )
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