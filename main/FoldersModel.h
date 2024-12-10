#ifndef FOLDERSMODEL_H
#define FOLDERSMODEL_H

#include <QString>
#include <QStandardItemModel>

#include <memory>
#include <list>

namespace Outlook
{
    class Folder;
    class Folders;
}

struct SCurrFolderInfo;

class CFoldersModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CFoldersModel( QObject *parent );

    virtual ~CFoldersModel();

    void reload();

    QString fullPathForItem( const QModelIndex &index ) const;
    QString fullPathForItem( QStandardItem *item ) const;

    std::shared_ptr< Outlook::Folder > folderForItem( const QModelIndex &index ) const;
    std::shared_ptr< Outlook::Folder > folderForItem( QStandardItem *item ) const;

    QString pathForItem( const QModelIndex &index ) const;
    QString pathForItem( QStandardItem *item ) const;

    void clear();

    QModelIndex addFolder( const QModelIndex &parentIndex, QWidget *parent );
    QModelIndex addFolder( const QModelIndex &parentIndex, const QString &folderName );

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedLoadingChildren( QStandardItem *parent );
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotReload();
    void slotAddNextFolder( QStandardItem *parent );

private:
    void addSubFolders( const std::shared_ptr< Outlook::Folder > &rootFolder );
    void addSubFolders( QStandardItem *item, const std::shared_ptr< Outlook::Folder > &parentFolder );

    [[nodiscard]] QStandardItem *addFolder( const std::shared_ptr< Outlook::Folder > &rootFolder, QStandardItem *parentItem );
    [[nodiscard]] QStandardItem *addFolder( const Outlook::Folder *rootFolder, QStandardItem *parentItem );

    std::unordered_map< QStandardItem *, std::unique_ptr< SCurrFolderInfo > > fFolders;
    std::unordered_map< QStandardItem *, std::shared_ptr< Outlook::Folder > > fFolderMap;

    int fCurrFolderNum{ 0 };
    int fNumFolders{ 0 };
};

#endif
