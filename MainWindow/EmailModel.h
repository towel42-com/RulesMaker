#ifndef EMAILMODEL_H
#define EMAILMODEL_H

#include <QString>
#include <QStandardItemModel>
#include <QVector>

#include <optional>
#include <memory>
#include <list>
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

    std::map< QString, CEmailAddressSection * > fChildItems;
};

class CEmailModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CEmailModel( QObject *parent );
    virtual ~CEmailModel();

    void reload();
    void clear();

    std::shared_ptr< Outlook::MailItem > emailItemFromIndex( const QModelIndex &idx ) const;
    std::shared_ptr< Outlook::MailItem > emailItemFromItem( QStandardItem * item ) const;

    QStringList rulesForIndex( const QModelIndex &idx ) const;
    QStringList rulesForItem( QStandardItem *item ) const;

    QString displayNameForIndex( const QModelIndex &idx ) const;
    QString displayNameForItem( QStandardItem *item ) const;

    void displayEmail( const QModelIndex &idx ) const;
    void displayEmail( QStandardItem *item ) const;

Q_SIGNALS:
    void sigFinishedGrouping();
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotGroupNextMailItemBySender();

private:
    void sortAll( QStandardItem *root );
    QString ruleForItem( QStandardItem *item ) const;
    void addEmailAddress( std::shared_ptr< Outlook::MailItem > mailItem );
    CEmailAddressSection *findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent, const QString &displayName );

    void addToDisplayName( CEmailAddressSection *currItem, const QString &displayName );

    std::shared_ptr< Outlook::Items > fItems{ nullptr };
    mutable std::optional< int > fItemCountCache;

    std::map< QString, CEmailAddressSection * > fRootItems;
    std::map< QString, CEmailAddressSection * > fCache;
    std::map< QString, CEmailAddressSection * > fDomainCache;
    std::map< QStandardItem *, std::shared_ptr< Outlook::MailItem > > fEmailCache;
    int fCurrPos{ 1 };
};

#endif
