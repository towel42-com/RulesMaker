#ifndef FOLDERSMODEL_H
#define FOLDERSMODEL_H

#include <QString>
#include <QStandardItemModel>

#include <memory>
#include <list>

class QProgressDialog;
namespace Outlook
{
    class MAPIFolder;
    class Folders;
}

class CFoldersModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CFoldersModel( QObject *parent );

    virtual ~CFoldersModel();

    void reload();

    QString fullPath( const QModelIndex &index ) const;
    QString fullPath( QStandardItem *item ) const;

    QString currentPath( const QModelIndex &index ) const;
    QString currentPath( QStandardItem *item ) const;

    void clear();

    void addFolder( const QModelIndex &idx, QWidget *parent );

Q_SIGNALS:
    void sigFinishedLoading();

private slots:
    void slotReload();
    void slotAddFolder( Outlook::MAPIFolder *folder );
    void slotFolderChanged( Outlook::MAPIFolder *folder );

private:
    void addSubFolders( std::shared_ptr< Outlook::MAPIFolder > rootFolder );
    bool addSubFolders( QStandardItem *item, std::shared_ptr< Outlook::MAPIFolder > parentFolder, QProgressDialog *progress );   // returns true if progress cancelled
    std::unordered_map< QStandardItem *, std::shared_ptr< Outlook::MAPIFolder > > fFolderMap;
};

#endif
