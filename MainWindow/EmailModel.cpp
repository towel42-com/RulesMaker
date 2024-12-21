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

CEmailModel::~CEmailModel()
{
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
        addEmailAddress( mailItem );

    emit sigSetStatus( fCurrPos, limit );

    if ( fCurrPos == limit )
    {
        sortAll( nullptr );
        processChildDisplayName();
        dumpNodes();
        emit sigFinishedGrouping();
    }
    else
    {
        fCurrPos++;
        QTimer::singleShot( 0, this, &CEmailModel::slotGroupNextMailItemBySender );
    }
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

QString CEmailModel::displayNameForIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto mailItem = mailItemFromIndex( idx );
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

    auto emailItem = mailItemFromItem( item );
    if ( !emailItem )
        return;
    COutlookAPI::instance()->displayEmail( emailItem );
}

CEmailAddressSection *CEmailModel::item( int row, int column /* = 0 */ ) const
{
    return dynamic_cast< CEmailAddressSection * >( QStandardItemModel::item( row, column ) );
}

QStringList CEmailModel::matchTextForIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    return matchTextListForItem( item );
}

std::unordered_set< QStandardItem * > getLeafChildren( QStandardItem *item )
{
    if ( !item )
        return {};

    if ( item->column() != 0 )
    {
        auto parent = item->parent();
        if ( !parent )
        {
            parent = item->model()->invisibleRootItem();
        }
        if ( !parent )
            return {};

        auto sibling = parent->child( item->row(), 0 );
        if ( !sibling )
            return {};
        item = sibling;
    }

    if ( !item )
        return {};

    if ( item->rowCount() == 0 )
        return { item };

    std::unordered_set< QStandardItem * > retVal;
    for ( auto ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii );
        auto children = getLeafChildren( child );
        for ( auto &&ii : children )
        {
            retVal.insert( ii );
        }
    }
    return retVal;
}

QStringList CEmailModel::matchTextListForItem( QStandardItem *item ) const
{
    return matchTextListForItem( dynamic_cast< CEmailAddressSection * >( item ) );
}

QStringList CEmailModel::matchTextListForItem( CEmailAddressSection *item ) const
{
    if ( !item )
        return {};

    QStringList retVal;
    if ( item->needsDisplayName() )
        return { item->matchTextForItem() };

    if ( item->parent() )
    {
        auto matchText = item->matchTextForItem();
        if ( ( item->column() == 0 ) && ( matchText.indexOf( '@' ) == -1 ) )
            matchText = "@" + matchText;
        retVal.push_back( matchText );
    }

    QStringList childNames;
    for ( auto &&ii = 0; ii < item->rowCount(); ++ii )
    {
        auto child = item->child( ii, 0 );

        if ( child->needsDisplayName() )
            childNames << child->matchTextForItem();
        else
        {
            if ( !child || child->text().isEmpty() )
                continue;

            auto curr = matchTextListForItem( child );
            if ( !curr.isEmpty() )
            {
                childNames << curr;
            }
        }
    }

    if ( item->needsDisplayName( true ) )
        retVal = childNames;
    else
        retVal << childNames;

    retVal.removeDuplicates();
    retVal.removeAll( QString() );
    return retVal;
}

std::shared_ptr< Outlook::MailItem > CEmailModel::mailItemFromIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() )
        return {};
    auto item = itemFromIndex( idx );
    return mailItemFromItem( item );
}

std::shared_ptr< Outlook::MailItem > CEmailModel::mailItemFromItem( const QStandardItem *item ) const
{
    if ( !item )
        return {};
    auto pos = fEmailCache.find( item );
    if ( pos == fEmailCache.end() )
    {
        if ( ( item->column() != 1 ) && item->parent() )
        {
            auto sibling = item->parent()->child( item->row(), 1 );
            return mailItemFromItem( sibling );
        }
        return {};
    }
    return ( *pos ).second;
}

std::pair< CEmailAddressSection *, QList< QStandardItem * > > makeRow( const QString &section, bool inBack, const QString &displayName )
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

bool CEmailAddressSection::needsDisplayName( bool includeAllChildren ) const
{
    if ( !fItemNeedsDisplayName.has_value() )
    {
        fItemNeedsDisplayName = false;
        if ( parent() && ( rowCount() == 0 ) )
        {
            auto emailAddress = matchTextForItem( true );
            static std::list< QString > knownDomains = { "gmail.com", "shopifyemail.com", "mailchimpapp.com" };

            for ( auto &&jj : knownDomains )
            {
                if ( emailAddress.indexOf( jj ) != -1 )
                {
                    fItemNeedsDisplayName = true;
                    break;
                }
            }
        }
    }
    if ( includeAllChildren && fItemNeedsDisplayName.value() )
    {
        for ( int ii = 0; ii < rowCount(); ++ii )
        {
            auto curr = child( ii );
            if ( !curr )
                continue;
            if ( !curr->needsDisplayName( true ) )
                return false;
        }
    }

    return fItemNeedsDisplayName.value();
}

QString CEmailAddressSection::matchTextForItem( bool forceNoDisplayName ) const
{
    QStringList path;
    QString separator = ".";
    if ( column() )
    {
        auto colZeroItem = parent() ? parent()->child( row(), 0 ) : nullptr;
        if ( !colZeroItem )
            return {};
        return colZeroItem->matchTextForItem( forceNoDisplayName );
    }

    if ( !forceNoDisplayName )
    {
        if ( !needsDisplayName() )
        {
            auto model = dynamic_cast< CEmailModel * >( this->model() );
            if ( !model )
                return {};

            auto mailItem = model->mailItemFromItem( this );
            auto emailAddresses = COutlookAPI::getEmailAddresses( mailItem, COutlookAPI::EAddressTypes::eSender | COutlookAPI::EAddressTypes::eSMTPOnly );
            if ( emailAddresses.second.empty() )
                return {};
            return "From: " + emailAddresses.second.front();
        }
    }

    if ( !hasChildren() )
        separator = "@";

    auto itemText = text();
    if ( itemText.isEmpty() )
    {
        int column = ( this->column() == 0 ) ? 1 : 0;
        itemText = parent() ? parent()->child( row(), column )->text() : QString();
    }
    path.push_back( itemText );

    if ( parent() )
        path.push_back( parent()->matchTextForItem( forceNoDisplayName ) );

    return path.join( separator );
}

CEmailAddressSection *CEmailAddressSection::child( int row, int column /*= 0 */ ) const
{
    return dynamic_cast< CEmailAddressSection * >( QStandardItem::child( row, column ) );
}

CEmailAddressSection *CEmailAddressSection::parent() const
{
    return dynamic_cast< CEmailAddressSection * >( QStandardItem::parent() );
}

void CEmailModel::processChildDisplayName()
{
    auto count = rowCount();
    for ( int ii = 0; ii < count; ++ii )
    {
        auto child = item( ii );
        if ( !child )
            continue;
        child->processChildDisplayName();
    }
}

void CEmailAddressSection::processChildDisplayName()
{
    bool allChildrenNeedDisplayName = true;
    qDebug() << "processing " << text();

    if ( rowCount() == 0 )
    {
        // no children
        fAllChildrenNeedDisplayName = needsDisplayName( true );
    }
    else
    {
        for ( auto ii = 0; ii < rowCount(); ++ii )
        {
            auto child = this->child( ii );
            if ( !child )
                continue;
            fAllChildrenNeedDisplayName = child->needsDisplayName( true ) && fAllChildrenNeedDisplayName;
            child->processChildDisplayName();
        }
    }
}

void CEmailModel::dumpNodes() const
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
    qDebug() << QString( "%1%2 - %3 - %4" ).arg( QString( depth, ' ' ), text(), ( needsDisplayName( true ) ? "true" : "false" ), ( needsDisplayName( false ) ? "true" : "false" ) );

    for ( int ii = 0; ii < rowCount(); ++ii )
    {
        auto child = this->child( ii );
        child->dumpNodes( depth + 1 );
    }
}
