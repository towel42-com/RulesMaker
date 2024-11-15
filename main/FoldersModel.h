#ifndef FOLDERSMODEL_H
#define FOLDERSMODEL_H

#include <QString>
#include <QStandardItemModel>

#include <memory>
#include <list>

class QProgressDialog;
namespace Outlook
{
    class Folder;
    class Folders;
}

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

private slots:
    void slotReload();
    void slotAddFolder( Outlook::Folder *folder );
    void slotFolderChanged( Outlook::Folder *folder );

private:
    void addSubFolders( std::shared_ptr< Outlook::Folder > rootFolder );
    bool addSubFolders( QStandardItem *item, std::shared_ptr< Outlook::Folder > parentFolder, QProgressDialog *progress );   // returns true if progress cancelled
    std::unordered_map< QStandardItem *, std::shared_ptr< Outlook::Folder > > fFolderMap;
};

#endif
