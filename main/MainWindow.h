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
    void slotSelectServerAndInbox();
    void slotSelectServer();
    void slotReloadAll();
    void slotReloadEmail();
    void slotReloadFolders();
    void slotReloadRules();
    void slotAddRule();
    void slotRunRule();
    void slotAddToSelectedRule();
    void slotRenameRules();
    void slotMergeRules();
    void slotSortRules();
    void slotMoveFromToAddress();

protected:
    void clearViews();
    std::unique_ptr< Ui::CMainWindow > fImpl;
};

#endif
