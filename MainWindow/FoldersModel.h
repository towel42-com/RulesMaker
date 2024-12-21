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
    void reloadJunk();
    void reloadTrash();

    QString fullPathForIndex( const QModelIndex &index ) const;
    QString fullPathForItem( QStandardItem *item ) const;

    std::shared_ptr< Outlook::Folder > folderForIndex( const QModelIndex &index ) const;
    std::shared_ptr< Outlook::Folder > folderForItem( QStandardItem *item ) const;

    QStandardItem *itemForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const;
    QModelIndex indexForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const;

    QString pathForIndex( const QModelIndex &index ) const;
    QString pathForItem( QStandardItem *item ) const;

    void clear();

    QModelIndex addFolder( const QModelIndex &parentIndex, QWidget *parent );
    QModelIndex addFolder( const QModelIndex &parentIndex, const QString &folderName );

    QModelIndex inboxIndex() const;
Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedLoadingChildren( QStandardItem *parent );
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotReload();
    void slotLoadNextFolder( QStandardItem *parent );

private:
    void loadRootFolders( const std::list< std::shared_ptr< Outlook::Folder > > &rootFolder );
    void loadSubFolders( QStandardItem *item, const std::shared_ptr< Outlook::Folder > &parentFolder );

    void removeFolder( const std::shared_ptr< Outlook::Folder > &ii );

    void updateMaps( QStandardItem *child, const std::shared_ptr< Outlook::Folder > &folder );

    [[nodiscard]] QStandardItem *loadFolder( const std::shared_ptr< Outlook::Folder > &folder, QStandardItem *parentItem );


    std::unordered_map< QStandardItem *, std::unique_ptr< SCurrFolderInfo > > fFolders;
    std::unordered_map< QStandardItem *, std::shared_ptr< Outlook::Folder > > fItemToFolderMap;
    std::map< QString, QStandardItem * > fFolderToItemMap;

    int fCurrFolderNum{ 0 };
    int fNumFolders{ 0 };
};

#endif
