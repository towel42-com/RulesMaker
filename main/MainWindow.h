#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QWidget>
#include <memory>
namespace Ui
{
    class CMainWindow;
}

class QModelIndex;
class CEmailModel;

class CMainWindow : public QWidget
{
    Q_OBJECT

public:
    explicit CMainWindow( QWidget *parent = nullptr );
    ~CMainWindow();

protected slots:
    void slotReload();

protected:
    std::unique_ptr< Ui::CMainWindow > fImpl;
};

#endif
