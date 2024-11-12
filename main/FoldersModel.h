#ifndef FOLDERSMODEL_H
#define FOLDERSMODEL_H

#include <QString>
#include <QStandardItemModel>

#include <memory>
#include <list>

namespace Outlook
{
    class MAPIFolder;
    class Folders;
}

class CFoldersModel : public QStandardItemModel
{
public:
    explicit CFoldersModel( QObject *parent );
    virtual ~CFoldersModel();

    //int rowCount( const QModelIndex &parent = QModelIndex() ) const;
    //int columnCount( const QModelIndex &parent ) const;
    //QVariant headerData( int section, Qt::Orientation orientation, int role ) const;
    //QVariant data( const QModelIndex &index, int role ) const;

    void changeItem( const QModelIndex &index, const QString &folderName );
    void addItem( const QString &folderName );
    void update();

    QString fullPath( const QModelIndex &index ) const;
    QString fullPath( QStandardItem *item ) const;

private:
    void addSubFolders( QStandardItem *item, std::shared_ptr< Outlook::MAPIFolder > parentFolder );
};

#endif
