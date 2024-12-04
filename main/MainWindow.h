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
class CWidgetWithStatus;

class CMainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit CMainWindow( QWidget *parent = nullptr );

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
    void slotUpdateActions();

    void slotStatusMessage( const QString & msg );
    void slotSetStatus( const QString &label, int curr, int max );
    void slotInitStatus( const QString &label, int max );
    void slotIncStatusValue( const QString &label );;
    void slotFinishedStatus( const QString &label );

protected:
    CStatusProgress *getProgressBar( const QString &label );
    void setupStatusBar();

    CStatusProgress *addStatusBar( QString label, CWidgetWithStatus *object );

    void clearViews();
    std::unique_ptr< Ui::CMainWindow > fImpl;
    std::map< QString, CStatusProgress * > fProgressBars;
    QPushButton *fCancelButton{ nullptr };
};

#endif
