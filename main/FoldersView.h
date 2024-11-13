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
Q_SIGNALS:
    void sigFinishedLoading();

protected slots:
    void itemSelected( const QModelIndex &index );

protected:
    std::shared_ptr< CFoldersModel > fModel;
    std::unique_ptr< Ui::CFoldersView > fImpl;
};

#endif   // CFoldersView_H
