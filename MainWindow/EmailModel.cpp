#include "EmailModel.h"
#include "OutlookAPI/OutlookAPI.h"

#include <QTimer>
#include <algorithm>
#include <QDebug>

#ifdef _DEBUG
// #define LIMIT_EMAIL_READ
#endif
CEmailModel::CEmailModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    clear();
    connect( COutlookAPI::instance().get(), &COutlookAPI::sigOptionChanged, this, &CEmailModel::reload );
}

void CEmailModel::clear()
{
    QStandardItemModel::clear();
    setHorizontalHeaderLabels(
        QStringList() << "Domain"
                      << "Sender"
                      << "Display Name" );
    beginResetModel();
    fItems.reset();
    fRootItems.clear();
    fCache.clear();
    fDomainCache.clear();
    fEmailCache.clear();
    fItemCountCache.reset();
    fCurrPos = 1;
    endResetModel();
}

void CEmailModel::reload()
{
    COutlookAPI::instance()->slotClearCanceled();

    beginResetModel();
    clear();

    std::tie( fItems, fItemCountCache ) = COutlookAPI::instance()->getEmailItemsForRootFolder();

    endResetModel();
    fCurrPos = 1;
    QTimer::singleShot( 0, this, &CEmailModel::slotGroupNextMailItemBySender );
}

void CEmailModel::slotGroupNextMailItemBySender()
{
    if ( !fItems || !fItemCountCache.has_value() )
        return;

    if ( fCurrPos > fItemCountCache.value() )
        return;

    if ( COutlookAPI::instance()->canceled() )
    {
        clear();
        emit sigFinishedGrouping();
        return;
    }

    auto mailItem = COutlookAPI::instance()->getEmailItem( fItems, fCurrPos );
    if ( mailItem )
        addEmailAddress( mailItem );

    emit sigSetStatus( fCurrPos, fItemCountCache.value() );

#ifdef LIMIT_EMAIL_READ
    if ( fCurrPos >= 100 )
        return;
#endif
    if ( fCurrPos == fItemCountCache.value() )
    {
        sortAll( nullptr );
        emit sigFinishedGrouping();
    }
    else
    {
        fCurrPos++;
        QTimer::singleShot( 0, this, &CEmailModel::slotGroupNextMailItemBySender );
    }
}

CEmailModel::~CEmailModel()
{
}

void CEmailModel::sortAll( QStandardItem *root )
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

void CEmailModel::addEmailAddress( std::shared_ptr< Outlook::MailItem > mailItem )
{
    if ( !mailItem )
        return;

    auto emailAddresses = COutlookAPI::getEmailAddresses( mailItem, COutlookAPI::EAddressTypes::eSender | COutlookAPI::EAddressTypes::eSMTPOnly );
    for ( int ii = 0; ( ii < emailAddresses.first.length() ) && ( ii < emailAddresses.second.length() ); ++ii )
    {
        auto emailAddress = emailAddresses.first[ ii ];
        auto displayName = emailAddresses.second[ ii ];
        auto key = emailAddress + "<" + displayName + ">"; 
        auto pos = fCache.find( key );
        if ( pos != fCache.end() )
        {
            continue;
        }

        qDebug() << "Processing Email: " << key;

        auto split = emailAddress.splitRef( '@', QString::SplitBehavior::SkipEmptyParts );
        if ( split.empty() )
            continue;

        auto user = split.front();
        QStringRef domain;
        if ( split.count() != 2 )
            continue;

        domain = split.back();
        pos = fDomainCache.find( domain.toString() );
        if ( pos != fDomainCache.end() )
        {
            auto retVal = findOrAddEmailAddressSection( user, {}, ( *pos ).second, displayName );
            fCache[ key ] = retVal;
            fEmailCache[ retVal ] = mailItem;
            continue;
        }

        auto list = domain.split( '.', QString::SplitBehavior::SkipEmptyParts );
        std::reverse( std::begin( list ), std::end( list ) );
        list.push_back( user );
        auto retVal = findOrAddEmailAddressSection( list.front(), list.mid( 1 ), nullptr, displayName );
        if ( retVal )
        {
            fCache[ key ] = retVal;
            fEmailCache[ retVal ] = mailItem;
            auto parent = dynamic_cast< CEmailAddressSection * >( retVal->parent() );
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


QString CEmailModel::displayNameForIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto mailItem = emailItemFromIndex( idx );
    if ( !mailItem )
        return {};
    auto emailAddresses = COutlookAPI::getEmailAddresses( mailItem, COutlookAPI::EAddressTypes::eSender | COutlookAPI::EAddressTypes::eSMTPOnly ).second;
    if ( emailAddresses.isEmpty() )
        return {};
    return emailAddresses.front();
}

QString CEmailModel::displayNameForItem( QStandardItem *item ) const
{
    return displayNameForIndex( indexFromItem( item ) );
}

void CEmailModel::displayEmail( const QModelIndex &idx ) const
{
    auto item = itemFromIndex( idx );
    displayEmail( item );
}

void CEmailModel::displayEmail( QStandardItem *item ) const
{
    if ( !item )
        return;

    auto emailItem = emailItemFromItem( item );
    if ( !emailItem )
        return;
    COutlookAPI::instance()->displayEmail( emailItem );
}

QStringList CEmailModel::matchTextForIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    return matchTextListForItem( item );
}

QStringList CEmailModel::matchTextListForItem( QStandardItem *item ) const
{
    QStringList retVal;

    if ( item->parent() )
    {
        auto matchText = matchTextForItem( item );
        if ( ( item->column() == 0 ) && ( matchText.indexOf( '@' ) == -1 ) )
            matchText = "@" + matchText;
        retVal.push_back( matchText );
    }
    for ( auto &&ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii, 0 );
        if ( !child || child->text().isEmpty() )
            continue;

        auto curr = matchTextListForItem( child );
        if ( !curr.isEmpty() )
            retVal << curr;
    }
    retVal.removeDuplicates();
    retVal.removeAll( QString() );
    return retVal;
}

QString CEmailModel::matchTextForItem( QStandardItem *item ) const
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

std::shared_ptr< Outlook::MailItem > CEmailModel::emailItemFromIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    return emailItemFromItem( item );
}

std::shared_ptr< Outlook::MailItem > CEmailModel::emailItemFromItem( QStandardItem * item ) const 
{
    if ( !item )
        return {};
    auto pos = fEmailCache.find( item );
    if ( pos == fEmailCache.end() )
    {
        if ( ( item->column() != 1 ) && item->parent() )
        {
            auto sibling = item->parent()->child( item->row(), 1 );
            return emailItemFromItem( sibling );
        }
        return {};
    }
    return ( *pos ).second;
}

std::pair< CEmailAddressSection *, QList< QStandardItem * > > makeRow( const QString &section, bool inBack, const QString & displayName )
{
    auto item = new CEmailAddressSection( section );

    QList< QStandardItem * > row;
    if ( inBack )
        row << new CEmailAddressSection << item;
    else
        row << item << new CEmailAddressSection;

    if ( inBack && !displayName.isEmpty() )
        row << new CEmailAddressSection( displayName );
    return { item, row };
}

CEmailAddressSection *CEmailModel::findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent, const QString &displayName )
{
    CEmailAddressSection *retVal{ nullptr };
    auto key = curr.toString().toLower();

    if ( parent )
    {
        auto pos = parent->fChildItems.find( key );
        if ( pos == parent->fChildItems.end() )
        {
            auto &&[ item, nextRemaining ] = makeRow( key, remaining.empty(), displayName );

            parent->appendRow( nextRemaining );
            parent->fChildItems[ key ] = item;
            retVal = item;
        }
        else
        {
            retVal = ( *pos ).second;
            if ( retVal && remaining.empty() )
            {
                addToDisplayName( retVal, displayName );
            }
        }
    }
    else
    {
        auto pos = fRootItems.find( key );
        if ( pos == fRootItems.end() )
        {
            auto &&[ item, nextRemaining ] = makeRow( key, remaining.empty(), displayName );
            appendRow( nextRemaining );
            fRootItems[ key ] = item;
            retVal = item;
        }
        else
        {
            retVal = ( *pos ).second;
            if ( retVal && remaining.empty() )
            {
                addToDisplayName( retVal, displayName );
            }
        }
    }

    if ( !remaining.empty() )
        return findOrAddEmailAddressSection( remaining.front(), remaining.mid( 1 ), retVal, displayName );
    return retVal;
}

void CEmailModel::addToDisplayName( CEmailAddressSection *retVal, const QString &displayName )
{
    if ( retVal->rowCount() >= 1 )
    {
        auto dispNameItem = retVal->child( 0, 2 );
        if ( dispNameItem )
        {
            auto text = dispNameItem->text() += ";" + displayName;
            dispNameItem->setText( text );
        }
        else
        {
            retVal->setChild( 0, 2, new CEmailAddressSection( displayName ) );
        }
    }
}
