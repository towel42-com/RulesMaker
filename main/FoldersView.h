#ifndef CFoldersVIEW_H
#define CFoldersVIEW_H

#include <QWidget>
#include <memory>
namespace Ui
{
    class CFoldersView;
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

    void reload();
    void clear();

    QString currentPath() const;
    QString fullPath() const;
Q_SIGNALS:
    void sigFinishedLoading();
    void sigFolderSelected( const QString &folderPath );

protected slots:
    void slotItemSelected( const QModelIndex &index );
    void slotAddFolder();

protected:
    std::shared_ptr< CFoldersModel > fModel;
    std::unique_ptr< Ui::CFoldersView > fImpl;
};

#endif   // CFoldersView_H
