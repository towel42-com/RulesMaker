#ifndef FOLDERSDLG_H
#define FOLDERSDLG_H

#include <QDialog>
#include <memory>
namespace Ui
{
    class CFoldersDlg;
}

namespace Outlook
{
    class Folder;
}

class QModelIndex;
class CFoldersModel;

class CFoldersDlg : public QDialog
{
    Q_OBJECT

public:
    explicit CFoldersDlg( QWidget *parent = nullptr );
    ~CFoldersDlg();

    QString currentPath() const;
    QString fullPath() const;
    std::shared_ptr< Outlook::Folder > selectedFolder();
Q_SIGNALS:

protected slots:
protected:
    std::unique_ptr< Ui::CFoldersDlg > fImpl;
};

#endif   // CFoldersView_H
