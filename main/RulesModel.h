#ifndef RULESMODEL_H
#define RULESMODEL_H

#include <QString>
#include <QAbstractListModel>

#include <memory>

namespace Outlook
{
    class Items;
}

class CRulesModel : public QAbstractListModel
{
public:
    explicit CRulesModel( QObject *parent );
    virtual ~CRulesModel();

    int rowCount( const QModelIndex &parent = QModelIndex() ) const;
    int columnCount( const QModelIndex &parent ) const;
    QVariant headerData( int section, Qt::Orientation orientation, int role ) const;
    QVariant data( const QModelIndex &index, int role ) const;

    void changeItem( const QModelIndex &index, const QString &folderName );
    void addItem( const QString &folderName );
    void update();

private:
    std::unique_ptr< Outlook::Items > fItems{ nullptr };

    mutable QHash< QModelIndex, QStringList > fCache;
};

#endif
