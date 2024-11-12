#ifndef CONTACTSMODEL_H
#define CONTACTSMODEL_H

#include <QString>
#include <QAbstractListModel>
#include <memory>
#include <optional>
namespace Outlook
{
    class Items;
}

class CContactsModel : public QAbstractListModel
{
public:
    explicit CContactsModel( QObject *parent );
    virtual ~CContactsModel();

    int rowCount( const QModelIndex &parent = QModelIndex() ) const;
    int columnCount( const QModelIndex &parent ) const;
    QVariant headerData( int section, Qt::Orientation orientation, int role ) const;
    QVariant data( const QModelIndex &index, int role ) const;

    void update();
    void changeItem( const QModelIndex &index, const QString &firstName, const QString &lastName, const QString &address, const QString &email );
    void addItem( const QString &firstName, const QString &lastName, const QString &address, const QString &email );

private:

    std::unique_ptr< Outlook::Items > fItems{ nullptr };

    mutable QHash< int, QStringList > fCache;
    mutable std::optional< int > fCountCache;
};

#endif
