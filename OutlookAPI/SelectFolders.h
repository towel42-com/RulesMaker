#ifndef SelectFolders_H
#define SelectFolders_H

#include <memory>
#include <optional>

#include <QDialog>
#include <QString>

namespace Outlook
{
    class Folder;
}

namespace Ui
{
    class CSelectFolders;
}

class CFoldersModel;
class CListFilterModel;

class CSelectFolders : public QDialog
{
    Q_OBJECT

public:
    explicit CSelectFolders( QWidget *parent = nullptr );
    ~CSelectFolders();

    void setFolders( const std::list< std::shared_ptr< Outlook::Folder > > &folders );
    std::list< std::shared_ptr< Outlook::Folder > > selectedFolders() const;
    virtual void accept() override;

Q_SIGNALS:

protected Q_SLOTS:
    void slotItemDoubleClicked( const QModelIndex &idx );


protected:
    void init();

    CFoldersModel *fModel{ nullptr };
    CListFilterModel *fFilterModel{ nullptr };

    std::unique_ptr< Ui::CSelectFolders > fImpl;
};

#endif   // CONTACTSVIEW_H
