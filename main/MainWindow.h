#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <memory>
namespace Ui
{
    class CMainWindow;
}

class QModelIndex;

class CMainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit CMainWindow( QWidget *parent = nullptr );

    void slotUpdateActions();

    ~CMainWindow();

protected slots:
    void slotSelectServer();
    void slotReload();

    void clearViews();

protected:
    std::unique_ptr< Ui::CMainWindow > fImpl;
};

#endif
