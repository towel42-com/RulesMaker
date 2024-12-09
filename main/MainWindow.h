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
    bool running() const;

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
    void slotEnableAllRules();

    void slotHandleProgressToggle();
    void slotUpdateActions();

    void slotAddFolderForSelectedEmail();

    void slotStatusMessage( const QString &msg );
    void slotSetStatus( const QString &label, int curr, int max );
    void slotInitStatus( const QString &label, int max );
    void slotIncStatusValue( const QString &label );
    ;
    void slotFinishedStatus( const QString &label );

protected:
    void updateActions();

    template< typename T >
    void setEnabled( T *item )
    {
        setEnabled( item, true, QStringList() );
    }
    template< typename T >
    void setEnabled( T *item, bool enabled, const QString &reason )
    {
        setEnabled( item, enabled, QStringList() << reason );
    }
    template< typename T >
    void setEnabled( T *item, bool enabled, QStringList reasons )
    {
        bool isRunning = running();
        if ( isRunning )
        {
            reasons = QStringList() << "Cannot execute while currently running";
            enabled = false;
        }

        item->setEnabled( enabled );
        if ( enabled )
            item->setToolTip( item->text() );
        else
        {
            QString msg;
            if ( reasons.size() == 1 )
            {
                msg = item->text() + " - " + reasons.front();
            }
            else
            {
                for ( auto &&ii : reasons )
                {
                    ii = "<li>" + ii + "<li>";
                }
                msg = reasons.join( "\n" );
                msg = item->text() + ":<ul>\n" + msg + "</ul>";
            }
            item->setToolTip( msg );
        }
    }
    CStatusProgress *getProgressBar( const QString &label );
    void setupStatusBar();

    CStatusProgress *addStatusBar( QString label, CWidgetWithStatus *object );

    void clearViews();
    std::unique_ptr< Ui::CMainWindow > fImpl;
    std::map< QString, CStatusProgress * > fProgressBars;
    QPushButton *fCancelButton{ nullptr };
};

#endif
