#include "FilterFromEmailModel.h"
#include "OutlookAPI/OutlookAPI.h"
#include "OutlookAPI/EmailAddress.h"

#include <QTimer>
#include <algorithm>
#include <QDebug>

#ifdef _DEBUG
// #define LIMIT_EMAIL_READ
#endif
CFilterFromEmailModel::CFilterFromEmailModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
    connect( COutlookAPI::instance().get(), &COutlookAPI::sigOptionChanged, this, &CFilterFromEmailModel::reload );
}

CFilterFromEmailModel::~CFilterFromEmailModel()
{
}

void CFilterFromEmailModel::clear()
{
    QStandardItemModel::clear();
    setHorizontalHeaderLabels(
        QStringList() << "Domain"
                      << "Sender"
                      << "From"
                      << "Subject" );
    beginResetModel();
    fItems.reset();
    fRootItems.clear();
    fDisplayNameEmailCache.clear();
    fDomainCache.clear();
    fEmailCache.clear();
    fItemCountCache.reset();
    fNumEmailsProcessed = 0;
    fNumEmailAddressesProcessed = 0;
    fUniqueEmails = 0;
    fCurrPos = 1;
    endResetModel();
}

void CFilterFromEmailModel::reload()
{
    COutlookAPI::instance()->slotClearCanceled();

    beginResetModel();
    clear();

    std::tie( fItems, fItemCountCache ) = COutlookAPI::instance()->getEmailItemsForRootFolder();

    endResetModel();
    fCurrPos = 1;
    QTimer::singleShot( 0, this, &CFilterFromEmailModel::slotGroupNextMailItemBySender );
}

void CFilterFromEmailModel::slotGroupNextMailItemBySender()
{
    if ( !fItems || !fItemCountCache.has_value() )
        return;

    auto limit = fItemCountCache.value();
    if ( COutlookAPI::instance()->onlyProcessTheFirst500Emails() )
        limit = std::min( limit, 500 );
#ifdef LIMIT_EMAIL_READ
    limit = std::min( limit, 100 );
#endif

    if ( fCurrPos > limit )
        return;

    if ( COutlookAPI::instance()->canceled() )
    {
        clear();
        emit sigFinishedGrouping();
        return;
    }

    auto mailItem = COutlookAPI::instance()->getEmailItem( fItems, fCurrPos );
    if ( mailItem )
        addMailItem( mailItem );

    emit sigSetStatus( fCurrPos, limit );

    if ( fCurrPos == limit )
    {
        sortAll( nullptr );
        //dumpNodes();
        emit sigFinishedGrouping();
    }
    else
    {
        fCurrPos++;
        QTimer::singleShot( 0, this, &CFilterFromEmailModel::slotGroupNextMailItemBySender );
    }
}

void CFilterFromEmailModel::sortAll( QStandardItem *root )
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

void CFilterFromEmailModel::addMailItem( std::shared_ptr< Outlook::MailItem > mailItem )
{
    if ( !mailItem )
        return;

    auto subject = COutlookAPI::getSubject( mailItem );

    fNumEmailsProcessed++;

    auto emailAddresses = COutlookAPI::getEmailAddresses( mailItem, COutlookAPI::EAddressTypes::eSender );
    fNumEmailAddressesProcessed++;
    for ( auto && emailAddress : emailAddresses )
    {
        auto key = emailAddress->toString();
        auto pos = fDisplayNameEmailCache.find( key );
        if ( pos != fDisplayNameEmailCache.end() )
        {
            continue;
        }

        qDebug() << "Processing Email Address: " << key;

        auto split = emailAddress->emailAddress().splitRef( '@', Qt::SplitBehaviorFlags::SkipEmptyParts );
        if ( split.empty() )
            continue;

        auto user = split.front();
        QStringRef domain;
        if ( split.count() != 2 )
            continue;

        domain = split.back();
        auto list = domain.split( '.', Qt::SplitBehaviorFlags::SkipEmptyParts );
        std::reverse( std::begin( list ), std::end( list ) );
        list.push_back( user );
        auto retVal = findOrAddEmailAddressSection( list.front().toString(), list.mid( 1 ), nullptr, emailAddress->displayName(), subject );
        if ( retVal )
        {
            fDisplayNameEmailCache[ key ] = retVal;
            fEmailCache[ retVal ] = mailItem;
            auto parent = retVal->parent();
            if ( parent )
                fDomainCache[ domain.toString() ] = parent;
        }
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

CEmailAddressSection *CEmailAddressSection::getSibling( int columnNumber )
{
    return const_cast< CEmailAddressSection * >( const_cast< const CEmailAddressSection * >( this )->getSibling( columnNumber ) );
}

const CEmailAddressSection *CEmailAddressSection::getSibling( int columnNumber ) const
{
    if ( !this )
        return nullptr;

    if ( column() == columnNumber )
        return this;

    QStandardItem *parentItem = this->parent();
    if ( !parentItem )
    {
        parentItem = model()->invisibleRootItem();
    }
    if ( !parentItem )
        return {};

    auto sibling = parentItem->child( row(), columnNumber );
    if ( !sibling )
        return {};
    return dynamic_cast< CEmailAddressSection * >( sibling );
}

std::unordered_set< const CEmailAddressSection * > CEmailAddressSection::getLeafChildren() const
{
    auto item = getSibling( 0 );

    if ( !item )
        return {};

    if ( item->rowCount() == 0 )
        return { item };

    std::unordered_set< const CEmailAddressSection * > retVal;
    for ( auto ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii );
        auto children = child->getLeafChildren();
        for ( auto &&ii : children )
        {
            retVal.insert( ii );
        }
    }
    return retVal;
}

QStringList CFilterFromEmailModel::subjectsForIndex( const QModelIndex &idx, bool allChildren ) const
{
    if ( !idx.isValid() )
        return {};
    return subjectsForItem( itemFromIndex( idx ), allChildren );
}

QStringList CFilterFromEmailModel::subjectsForItem( QStandardItem *item, bool allChildren ) const
{
    return subjectsForItem( dynamic_cast< CEmailAddressSection * >( item ), allChildren );
}

QStringList CFilterFromEmailModel::subjectsForItem( const CEmailAddressSection *item, bool allChildren ) const
{
    std::unordered_set< const CEmailAddressSection * > items;
    if ( !allChildren )
        items.insert( item );
    else
    {
        items = item->getLeafChildren();
    }

    QStringList retVal;
    for ( auto &&ii : items )
    {
        auto mailItem = mailItemFromItem( ii );
        if ( !mailItem )
            continue;
        auto subject = COutlookAPI::getSubject( mailItem );
        if ( subject.isEmpty() )
            continue;

        retVal = mergeStringLists( retVal, { subject }, true );
    }
    return retVal;
}

QStringList CFilterFromEmailModel::sendersForIndex( const QModelIndex &idx, bool allChildren /*= false*/ ) const
{
    if ( !idx.isValid() )
        return {};
    return sendersForItem( itemFromIndex( idx ), allChildren );
}

QStringList CFilterFromEmailModel::sendersForItem( QStandardItem *item, bool allChildren /*= false*/ ) const
{
    return sendersForItem( dynamic_cast< CEmailAddressSection * >( item ), allChildren );
}

QStringList CFilterFromEmailModel::sendersForItem( const CEmailAddressSection *item, bool allChildren /*= false */ ) const
{
    std::unordered_set< const CEmailAddressSection * > items;
    if ( !allChildren )
        items.insert( item );
    else
    {
        items = item->getLeafChildren();
    }

    QStringList retVal;
    for ( auto &&ii : items )
    {
        auto mailItem = mailItemFromItem( ii );
        if ( !mailItem )
            continue;
        auto emailAddresses = COutlookAPI::getEmailAddresses( mailItem, COutlookAPI::EAddressTypes::eSender, COutlookAPI::EContactTypes::eAllContacts );
        if ( emailAddresses.empty() )
            continue;

        for ( auto &&ii : emailAddresses )
            retVal << ii->toString();

    }
    retVal.removeDuplicates();
    retVal.removeAll( QString() );

    return retVal;
}

QStringList CFilterFromEmailModel::displayNamesForIndex( const QModelIndex &idx, bool allChildren ) const
{
    if ( !idx.isValid() )
        return {};
    return displayNamesForItem( itemFromIndex( idx ), allChildren );
}

QStringList CFilterFromEmailModel::displayNamesForItem( QStandardItem *item, bool allChildren ) const
{
    return displayNamesForItem( dynamic_cast< CEmailAddressSection * >( item ), allChildren );
}

QStringList CFilterFromEmailModel::displayNamesForItem( const CEmailAddressSection *item, bool allChildren ) const
{
    std::unordered_set< const CEmailAddressSection * > items;
    if ( !allChildren )
        items.insert( item );
    else
    {
        items = item->getLeafChildren();
    }

    QStringList retVal;
    for ( auto &&ii : items )
    {
        auto mailItem = mailItemFromItem( ii );
        if ( !mailItem )
            continue;
        auto emailAddresses = COutlookAPI::getEmailAddresses( mailItem, COutlookAPI::EAddressTypes::eSender );
        if ( emailAddresses.empty() )
            continue;

        auto displayNames = getDisplayNames( emailAddresses );

        retVal = mergeStringLists( retVal, displayNames, true );
    }

    return retVal;
}

void CFilterFromEmailModel::displayEmail( const QModelIndex &idx ) const
{
    auto item = itemFromIndex( idx );
    displayEmail( item );
}

void CFilterFromEmailModel::displayEmail( QStandardItem *item ) const
{
    if ( !item )
        return;

    auto emailItem = mailItemFromItem( item );
    if ( !emailItem )
        return;
    COutlookAPI::instance()->displayEmail( emailItem );
}

CEmailAddressSection *CFilterFromEmailModel::item( int row, int column /* = 0 */ ) const
{
    return dynamic_cast< CEmailAddressSection * >( QStandardItemModel::item( row, column ) );
}

QString CFilterFromEmailModel::summary() const
{
    return QString( "%1 emails with %2 email addresses, %3 unique email addresses" ).arg( fNumEmailsProcessed ).arg( fNumEmailAddressesProcessed ).arg( fUniqueEmails );
}

QStringList CFilterFromEmailModel::matchTextForIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    return matchTextListForItem( item );
}

QStringList CFilterFromEmailModel::matchTextListForItem( QStandardItem *item ) const
{
    return matchTextListForItem( dynamic_cast< CEmailAddressSection * >( item ) );
}

QStringList CFilterFromEmailModel::matchTextListForItem( CEmailAddressSection *item ) const
{
    if ( !item )
        return {};

    QStringList retVal;

    if ( item->parent() )
    {
        auto matchText = matchTextForItem( item );
        if ( ( item->column() == 0 ) && ( matchText.indexOf( '@' ) == -1 ) )
            matchText = "@" + matchText;
        retVal.push_back( matchText );
    }
    retVal = mergeStringLists( retVal, {}, false );

    for ( auto &&ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii, 0 );
        if ( !child || child->text().isEmpty() )
            continue;

        auto curr = matchTextListForItem( child );
        retVal = mergeStringLists( retVal, curr, false );
    }

    return retVal;
}
QString CFilterFromEmailModel::matchTextForItem( QStandardItem *item ) const
{
    return matchTextForItem( dynamic_cast< CEmailAddressSection * >( item ) );
}

QString CFilterFromEmailModel::matchTextForItem( CEmailAddressSection *item ) const
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
        path.push_back( matchTextForItem( parent ) );

    return path.join( separator );
}

std::shared_ptr< Outlook::MailItem > CFilterFromEmailModel::mailItemFromIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    return mailItemFromItem( item );
}

std::shared_ptr< Outlook::MailItem > CFilterFromEmailModel::mailItemFromItem( const QStandardItem *item ) const
{
    return mailItemFromItem( dynamic_cast< const CEmailAddressSection * >( item ) );
}

std::shared_ptr< Outlook::MailItem > CFilterFromEmailModel::mailItemFromItem( const CEmailAddressSection *item ) const
{
    item = item->getSibling( 1 );

    if ( !item )
        return {};

    auto pos = fEmailCache.find( item );
    if ( pos == fEmailCache.end() )
    {
        //if ( ( item->column() != 1 ) && item->parent() )
        //{
        //    auto sibling = item->parent()->child( item->row(), 1 );
        //    return mailItemFromItem( sibling );
        //}
        return {};
    }
    return ( *pos ).second;
}

std::pair< CEmailAddressSection *, QList< QStandardItem * > > CFilterFromEmailModel::makeRow( const QString &section, bool inBack, const QString &displayName, const QString &subject )
{
    auto item = new CEmailAddressSection( section );

    QList< QStandardItem * > row;
    if ( !inBack )
        row << item << new CEmailAddressSection;
    else
    {
        row << new CEmailAddressSection << item;

        if ( !displayName.isEmpty() )
        {
            row << new CEmailAddressSection( displayName );
        }
        else
            row << new CEmailAddressSection;

        if ( !subject.isEmpty() )
        {
            row << new CEmailAddressSection( subject );
        }
        else
            row << new CEmailAddressSection;
        fUniqueEmails++;
    }

    return { item, row };
}

CEmailAddressSection *CFilterFromEmailModel::findOrAddEmailAddressSection( const QString &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent, const QString &displayName, const QString &subject )
{
    CEmailAddressSection *retVal{ nullptr };
    auto key = curr.toLower();

    if ( parent )
    {
        parent = parent->getSibling( 0 );
        auto pos = parent->fChildItems.find( key );
        if ( pos == parent->fChildItems.end() )
        {
            auto &&[ item, nextRemaining ] = makeRow( key, remaining.empty(), displayName, subject );

            parent->appendRow( nextRemaining );
            parent->fChildItems[ key ] = item;
            retVal = item;
        }
        else
        {
            retVal = ( *pos ).second;
        }
    }
    else
    {
        auto pos = fRootItems.find( key );
        if ( pos == fRootItems.end() )
        {
            auto &&[ item, nextRemaining ] = makeRow( key, remaining.empty(), displayName, subject );
            appendRow( nextRemaining );
            fRootItems[ key ] = item;
            retVal = item;
        }
        else
        {
            retVal = ( *pos ).second;
        }
    }

    if ( !remaining.empty() )
        return findOrAddEmailAddressSection( remaining.front().toString(), remaining.mid( 1 ), retVal, displayName, subject );
    return retVal;
}

CEmailAddressSection *CEmailAddressSection::child( int row, int column /*= 0 */ ) const
{
    return dynamic_cast< CEmailAddressSection * >( QStandardItem::child( row, column ) );
}

CEmailAddressSection *CEmailAddressSection::parent() const
{
    return dynamic_cast< CEmailAddressSection * >( QStandardItem::parent() );
}

void CFilterFromEmailModel::dumpNodes() const
{
    auto count = rowCount();
    for ( int ii = 0; ii < count; ++ii )
    {
        auto child = item( ii );
        if ( !child )
            continue;
        child->dumpNodes( 0 );
    }
}

void CEmailAddressSection::dumpNodes( int depth ) const
{
    //qDebug() << QString( "%1%2 - %3 - %4" ).arg( QString( depth, ' ' ), text() );

    for ( int row = 0; row < rowCount(); ++row )
    {
        QStringList text;
        for ( auto col = 0; col < columnCount(); ++col )
        {
            auto child = this->child( row, col );
            text << child->text();
        }
        qDebug() << QString( "%1%2" ).arg( QString( depth, ' ' ), text.join( ", " ) );
        this->child( row, 0 )->dumpNodes( depth + 1 );
    }
}
