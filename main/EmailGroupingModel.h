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
    class Items;
}

class CEmailAddressSection : public QStandardItem
{
public:
    CEmailAddressSection( const QString &itemName ) :
        QStandardItem( itemName )
    {
    }

    std::map< QString, CEmailAddressSection * > fChildItems;
};

class CEmailGroupingModel : public QStandardItemModel
{
public:
    explicit CEmailGroupingModel( QObject *parent );
    virtual ~CEmailGroupingModel();

    void addEmailAddress( const QString &email );

private:
    CEmailAddressSection *findOrAddEmailAddressSection( const QStringRef &curr, const QVector< QStringRef > &remaining, CEmailAddressSection *parent );

    std::shared_ptr< Outlook::Items > fEmailItems;

    std::map< QString, CEmailAddressSection * > fRootItems;
};

#endif
