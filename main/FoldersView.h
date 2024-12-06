#ifndef CFoldersVIEW_H
#define CFoldersVIEW_H

#include "WidgetWithStatus.h"
#include <memory>
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
class CListFilterModel;

class CFoldersView : public CWidgetWithStatus
{
    Q_OBJECT

public:
    explicit CFoldersView( QWidget *parent = nullptr );

    void init();

    ~CFoldersView();

    void reload( bool notifyOnFinished );
    void clear();
    void clearSelection();
    void addFolder( const QString &fileName );

    QString selectedPath() const;
    QString selectedFullPath() const;
    std::shared_ptr< Outlook::Folder > selectedFolder() const;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFolderSelected( const QString &folderPath );

protected Q_SLOTS:
    void slotItemSelected( const QModelIndex &index );
    void slotAddFolder();
    void slotSetRootFolder();

protected:
    void selectAndScroll( const QModelIndex &newIndex );
    QModelIndex currentIndex() const;
    QModelIndex sourceIndex( const QModelIndex &idx ) const;

    CFoldersModel *fModel{ nullptr };
    CListFilterModel *fFilterModel{ nullptr };
    std::unique_ptr< Ui::CFoldersView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CFoldersView_H
