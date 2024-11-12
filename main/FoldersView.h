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
    ~CFoldersView();

protected slots:
    void addEntry();
    void changeEntry();
    void itemSelected( const QModelIndex &index );

protected:
    std::shared_ptr< CFoldersModel > fModel;
    std::unique_ptr< Ui::CFoldersView > fImpl;
};

#endif   // CFoldersView_H
