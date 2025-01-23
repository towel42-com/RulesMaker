#ifndef FOLDERSMODEL_H
#define FOLDERSMODEL_H

#include "OutlookAPI/OutlookObj.h"

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

    COutlookObj< Outlook::MAPIFolder > folderForIndex( const QModelIndex &index ) const;
    COutlookObj< Outlook::MAPIFolder > folderForItem( QStandardItem *item ) const;

    QStandardItem *itemForFolder( const COutlookObj< Outlook::MAPIFolder > &folder ) const;
    QModelIndex indexForFolder( const COutlookObj< Outlook::MAPIFolder > &folder ) const;

    QString pathForIndex( const QModelIndex &index ) const;
    QString pathForItem( QStandardItem *item ) const;

    void clear();

    QModelIndex addFolder( const QModelIndex &parentIndex, QWidget *parent );
    QModelIndex addFolder( const QModelIndex &parentIndex, const QString &folderName );

    QModelIndex inboxIndex() const;

    QString summary() const;
Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedLoadingChildren( QStandardItem *parent );
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotReload();
    void slotLoadNextFolder( QStandardItem *parent );

private:
    void loadRootFolders( const std::list< COutlookObj< Outlook::MAPIFolder > > &rootFolder, bool setRootFolders );
    void loadSubFolders( QStandardItem *item, const COutlookObj< Outlook::MAPIFolder > &parentFolder );

    void removeFolder( const COutlookObj< Outlook::MAPIFolder > &ii, bool removeFromRootList = true );

    void updateMaps( QStandardItem *child, const COutlookObj< Outlook::MAPIFolder > &folder );

    [[nodiscard]] QStandardItem *loadFolder( const COutlookObj< Outlook::MAPIFolder > &folder, QStandardItem *parentItem );

    std::unordered_map< QStandardItem *, std::unique_ptr< SCurrFolderInfo > > fFolders;
    std::unordered_map< QStandardItem *, COutlookObj< Outlook::MAPIFolder > > fItemToFolderMap;
    std::map< QString, QStandardItem * > fFolderToItemMap;
    std::list< COutlookObj< Outlook::MAPIFolder > > fRootFolders;

    int fCurrFolderNum{ 0 };
    int fNumFolders{ 0 };
};

#endif
