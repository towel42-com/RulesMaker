#ifndef EMAILGROUPINGMODEL_H
#define EMAILGROUPINGMODEL_H

#include <QString>
#include <QStandardItemModel>
#include <QVector>

#include <memory>
#include <list>
#include <map>

namespace Outlook
{
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

class CEmailGroupingModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CEmailGroupingModel( QObject *parent );
    virtual ~CEmailGroupingModel();

    void addEmailAddress( std::shared_ptr< Outlook::MailItem >, const QString &email );

    void clear();

    std::shared_ptr< Outlook::MailItem > emailItemFromIndex( const QModelIndex &idx );

private:
    CEmailAddressSection *findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent );

    std::map< QString, CEmailAddressSection * > fRootItems;
    std::map< QString, CEmailAddressSection * > fCache;
    std::map< QString, CEmailAddressSection * > fDomainCache;
    std::map< QStandardItem *, std::shared_ptr< Outlook::MailItem > > fEmailCache;
};

#endif
