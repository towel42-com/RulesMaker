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

    void addFolder( const QModelIndex &idx, QWidget *parent );

Q_SIGNALS:
    void sigFinishedLoading();
    void sigIncStatusValue();
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotReload();
    void slotAddFolder( Outlook::Folder *folder );
    void slotFolderChanged( Outlook::Folder *folder );
    void slotAddNextFolder( QStandardItem *parent );

private:
    void addSubFolders( const std::shared_ptr< Outlook::Folder > &rootFolder );
    void addSubFolders( QStandardItem *item, const std::shared_ptr< Outlook::Folder > &parentFolder, bool root );   // returns true if progress cancelled

    std::unordered_map< QStandardItem *, std::unique_ptr< SCurrFolderInfo > > fFolders;
    std::unordered_map< QStandardItem *, std::shared_ptr< Outlook::Folder > > fFolderMap;
};

#endif
