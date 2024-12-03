#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <memory>
namespace Ui
{
    class CMainWindow;
}

class QPushButton;
class CStatusProgress;
class QModelIndex;

class CMainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit CMainWindow( QWidget *parent = nullptr );

    void slotUpdateActions();

    ~CMainWindow();

protected Q_SLOTS:
    void slotSelectServer();
    void slotReloadAll();

    void updateWindowTitle();

    void slotReloadEmail();
    void slotReloadFolders();
    void slotReloadRules();
    void slotAddRule();

    void clearSelection();

    void slotRunSelectedRule();
    void slotAddToSelectedRule();
    void slotRenameRules();
    void slotMergeRules();
    void slotSortRules();
    void slotMoveFromToAddress();
    void slotRunAllRules();
    void slotHandleProgressToggle();

protected:
    void setupStatusBar();

    CStatusProgress * addStatusBar( const QString &label, QObject *object, bool hasInc );

    void clearViews();
    std::unique_ptr< Ui::CMainWindow > fImpl;
    std::map< QString, CStatusProgress * > fProgressBars;
    QPushButton *fCancelButton{ nullptr };
};

#endif
