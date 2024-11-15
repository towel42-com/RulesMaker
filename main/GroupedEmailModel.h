#ifndef EMAILGROUPINGMODEL_H
#define EMAILGROUPINGMODEL_H

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
    CEmailAddressSection(){};

    CEmailAddressSection( const QString &itemName ) :
        QStandardItem( itemName )
    {
    }

    std::map< QString, CEmailAddressSection * > fChildItems;
};

class CGroupedEmailModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CGroupedEmailModel( QObject *parent );
    virtual ~CGroupedEmailModel();

    void addEmailAddress( std::shared_ptr< Outlook::MailItem >, const QString &email );

    void reload();
    void clear();

    std::shared_ptr< Outlook::MailItem > emailItemFromIndex( const QModelIndex &idx ) const;
    QStringList rulesForIndex( const QModelIndex &idx ) const;
    QStringList rulesForItem( QStandardItem *item ) const;
    void setOnlyGroupUnread( bool value );
    bool onlyGroupUnread() const { return fOnlyGroupUnread; }
    Q_SIGNALS:
    void sigFinishedGrouping();

private:
    QString ruleForItem( QStandardItem *item ) const;
    void groupMailItemsBySender( QWidget *parent );
    CEmailAddressSection *findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent );

    std::shared_ptr< Outlook::Items > fItems{ nullptr };

    std::map< QString, CEmailAddressSection * > fRootItems;
    std::map< QString, CEmailAddressSection * > fCache;
    std::map< QString, CEmailAddressSection * > fDomainCache;
    std::map< QStandardItem *, std::shared_ptr< Outlook::MailItem > > fEmailCache;
    mutable std::optional< int > fCountCache;
    bool fOnlyGroupUnread{ true };
};

#endif