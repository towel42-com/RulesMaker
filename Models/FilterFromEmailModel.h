#ifndef EMAILMODEL_H
#define EMAILMODEL_H

#include <QString>
#include <QStandardItemModel>
#include <QVector>

#include <optional>
#include <memory>
#include <list>
#include <unordered_map>
#include <unordered_set>
#include <map>

namespace Outlook
{
    class Items;
    class MailItem;
}

class CEmailAddressSection : public QStandardItem
{
public:
    CEmailAddressSection() {};

    CEmailAddressSection( const QString &itemName ) :
        QStandardItem( itemName ),
        fName( itemName )
    {
    }

    void dumpNodes( int depth = 0 ) const;
    const CEmailAddressSection *getSibling( int columnNumber ) const;
    CEmailAddressSection *getSibling( int columnNumber );
    std::unordered_set< const CEmailAddressSection * > getLeafChildren() const;

    std::map< QString, CEmailAddressSection * > fChildItems;

    CEmailAddressSection *child( int row, int column = 0 ) const;
    CEmailAddressSection *parent() const;

private:
    QString fName;
    bool fAllChildrenNeedDisplayName{ false };
};

class CFilterFromEmailModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CFilterFromEmailModel( QObject *parent );
    virtual ~CFilterFromEmailModel();

    void reload();
    void clear();

    std::shared_ptr< Outlook::MailItem > mailItemFromIndex( const QModelIndex &idx ) const;
    std::shared_ptr< Outlook::MailItem > mailItemFromItem( const QStandardItem *item ) const;
    std::shared_ptr< Outlook::MailItem > mailItemFromItem( const CEmailAddressSection *item ) const;

    QStringList matchTextForIndex( const QModelIndex &idx ) const;
    QStringList matchTextListForItem( QStandardItem *item ) const;

    QStringList displayNamesForIndex( const QModelIndex &idx, bool allChildren = false ) const;
    QStringList displayNamesForItem( QStandardItem *item, bool allChildren = false ) const;
    QStringList displayNamesForItem( const CEmailAddressSection *item, bool allChildren = false ) const;

    QStringList sendersForIndex( const QModelIndex &idx, bool allChildren = false ) const;
    QStringList sendersForItem( QStandardItem *item, bool allChildren = false ) const;
    QStringList sendersForItem( const CEmailAddressSection *item, bool allChildren = false ) const;

    QStringList subjectsForIndex( const QModelIndex &idx, bool allChildren = false ) const;
    QStringList subjectsForItem( QStandardItem *item, bool allChildren = false ) const;
    QStringList subjectsForItem( const CEmailAddressSection *item, bool allChildren = false ) const;

    void displayEmail( const QModelIndex &idx ) const;
    void displayEmail( QStandardItem *item ) const;

    CEmailAddressSection *item( int row, int column = 0 ) const;

    QString summary() const;
Q_SIGNALS:
    void sigFinishedGrouping();
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotGroupNextMailItemBySender();

private:
    void dumpNodes() const;

    QString matchTextForItem( QStandardItem *item ) const;
    QString matchTextForItem( CEmailAddressSection *item ) const;
    QStringList matchTextListForItem( CEmailAddressSection *item ) const;

    void sortAll( QStandardItem *root );
    void addMailItem( std::shared_ptr< Outlook::MailItem > mailItem );
    CEmailAddressSection *findOrAddEmailAddressSection( const QString &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent, const QString &displayName, const QString &subject );
    std::pair< CEmailAddressSection *, QList< QStandardItem * > > makeRow( const QString &section, bool inBack, const QString &displayName, const QString &subject );

    std::shared_ptr< Outlook::Items > fItems{ nullptr };
    mutable std::optional< int > fItemCountCache;

    int fNumEmailsProcessed{ 0 };
    int fNumEmailAddressesProcessed{ 0 };
    int fUniqueEmails{ 0 };
    std::map< QString, CEmailAddressSection * > fRootItems;
    std::map< QString, CEmailAddressSection * > fDisplayNameEmailCache;
    std::map< QString, CEmailAddressSection * > fDomainCache;
    std::map< const QStandardItem *, std::shared_ptr< Outlook::MailItem > > fEmailCache;
    int fCurrPos{ 1 };
};

#endif
