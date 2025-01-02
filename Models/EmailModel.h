#ifndef EMAILMODEL_H
#define EMAILMODEL_H

#include <QString>
#include <QStandardItemModel>
#include <QVector>

#include <optional>
#include <memory>
#include <list>
#include <unordered_map>
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
        QStandardItem( itemName )
    {
    }

    void dumpNodes( int depth = 0 ) const;

    std::map< QString, CEmailAddressSection * > fChildItems;

    CEmailAddressSection *child( int row, int column = 0 ) const;
    CEmailAddressSection *parent() const;

private:
    bool fAllChildrenNeedDisplayName{ false };
};

class CEmailModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CEmailModel( QObject *parent );
    virtual ~CEmailModel();

    void reload();
    void clear();

    std::shared_ptr< Outlook::MailItem > mailItemFromIndex( const QModelIndex &idx ) const;
    std::shared_ptr< Outlook::MailItem > mailItemFromItem( const QStandardItem *item ) const;

    QStringList matchTextForIndex( const QModelIndex &idx ) const;
    QStringList matchTextListForItem( QStandardItem *item ) const;

    QString displayNameForIndex( const QModelIndex &idx ) const;
    QString displayNameForItem( QStandardItem *item ) const;

    void displayEmail( const QModelIndex &idx ) const;
    void displayEmail( QStandardItem *item ) const;

    CEmailAddressSection *item( int row, int column = 0 ) const;
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
    void addEmailAddress( std::shared_ptr< Outlook::MailItem > mailItem );
    CEmailAddressSection *findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent, const QString &displayName );

    void addToDisplayName( CEmailAddressSection *currItem, const QString &displayName );

    std::shared_ptr< Outlook::Items > fItems{ nullptr };
    mutable std::optional< int > fItemCountCache;

    std::map< QString, CEmailAddressSection * > fRootItems;
    std::map< QString, CEmailAddressSection * > fCache;
    std::map< QString, CEmailAddressSection * > fDomainCache;
    std::map< const QStandardItem *, std::shared_ptr< Outlook::MailItem > > fEmailCache;
    int fCurrPos{ 1 };
};

#endif
