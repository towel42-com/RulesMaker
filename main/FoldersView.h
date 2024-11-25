#ifndef CFoldersVIEW_H
#define CFoldersVIEW_H

#include <QWidget>
#include <memory>
#include "Wrappers.h"
namespace Ui
{
    class CFoldersView;
}

namespace Outlook
{
    class Folder;
}

class QModelIndex;
class CFoldersModel;

class CFoldersView : public QWidget
{
    Q_OBJECT

public:
    explicit CFoldersView( QWidget *parent = nullptr );

    void init();

    ~CFoldersView();

    void reload( bool notifyOnFinished );
    void clear();

    QString selectedPath() const;
    QString selectedFullPath() const;
    std::shared_ptr< Outlook::Folder > selectedFolder() const;
Q_SIGNALS:
    void sigFinishedLoading();
    void sigFolderSelected( const QString &folderPath );

protected slots:
    void slotItemSelected( const QModelIndex &index );
    void slotAddFolder();

protected:
    std::shared_ptr< CFoldersModel > fModel;
    std::unique_ptr< Ui::CFoldersView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CFoldersView_H
