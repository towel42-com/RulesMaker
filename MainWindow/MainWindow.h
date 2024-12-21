#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <memory>
#include <list>
#include <utility>
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

Q_SIGNALS:
    void sigRunningStateChanged( bool running );

protected Q_SLOTS:
    void slotSelectServer();
    void slotReloadAll();

    void updateWindowTitle();

    void clearSelection();

    void slotReloadEmail();
    void slotReloadFolders();
    void slotReloadRules();

    void slotAddRule();
    void slotAddToSelectedRule();

    void slotRenameRules();
    void slotMergeRules();
    void slotSortRules();
    void slotMoveFromToAddress();
    void slotEnableAllRules();

    void slotAddFolderForSelectedEmail();

    void slotRunSelectedRule();
    void slotRunAllRules();
    void slotRunAllRulesOnAllFolders();
    void slotRunAllRulesOnSelectedFolder();
    void slotRunSelectedRuleOnSelectedFolder();

    void slotEmptyTrash();
    void slotEmptyJunkFolder();

    void slotHandleProgressToggle();
    void slotUpdateActions();

    void slotStatusMessage( const QString &msg );
    void slotSetStatus( const QString &label, int curr, int max );
    void slotInitStatus( const QString &label, int max );
    void slotIncStatusValue( const QString &label );

    void slotFinishedStatus( const QString &label );
    void slotAbout();

    void slotOptionsChanged();
    void slotSettings();

protected:
    void updateActions();

    template< typename T >
    void setEnabled( T *item )
    {
        setEnabled( item, { true, "" } );
    }

    using TReason = std::pair< bool, QString >;
    using TReasons = std::list< std::pair< bool, QString > >;
    template< typename T >
    void setEnabled( T *item, bool enabled, const QString &reason )
    {
        setEnabled( item, { enabled, reason } );
    }
    template< typename T >
    void setEnabled( T *item, const TReason &reason )
    {
        setEnabled( item, TReasons( { reason } ) );
    }
    template< typename T >
    void setEnabled( T *item, TReasons reasons )
    {
        if ( running() )
        {
            reasons.clear();
            reasons.emplace_back( false, "Cannot execute while processing" );
        }
        bool enabled = true;
        for ( auto &&ii : reasons )
        {
            enabled = enabled && ii.first;
        }

        item->setEnabled( enabled );
        if ( enabled )
            item->setToolTip( item->text() );
        else
        {
            QString msg;
            if ( reasons.size() == 1 )
            {
                msg = item->text() + " - " + reasons.front().second;
            }
            else
            {
                for ( auto &&ii : reasons )
                {
                    if ( !ii.first )
                        msg += "<li>" + ii.second + "<li>\n";
                }
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
