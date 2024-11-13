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
    Q_OBJECT;

public:
    explicit CFoldersModel( QObject *parent );

    virtual ~CFoldersModel();

    void reload();
    QString fullPath( const QModelIndex &index ) const;
    QString fullPath( QStandardItem *item ) const;

Q_SIGNALS:
    void sigFinishedLoading();

private:
    void addSubFolders( QStandardItem *item, std::shared_ptr< Outlook::MAPIFolder > parentFolder );
    void addSubFolders( std::shared_ptr< Outlook::MAPIFolder > rootFolder );
};

#endif
