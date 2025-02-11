#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QLabel>

#include <memory>
#include <list>
#include <utility>
#include <optional>
namespace Ui
{
    class CMainWindow;
}

namespace Outlook
{
    class Rule;
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

    void setWaitCursor( bool wait );
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
    void slotEnableAllRules();
    void slotDeleteAllDisabledRules();
    void slotFindEmptyFolders();

    void slotAddFolderForSelectedEmail();

    void slotRunSelectedRule();
    void slotRunAllRules();
    void slotRunAllRulesOnTrashFolder();
    void slotRunAllRulesOnJunkFolder();
    void slotRunAllRulesOnSelectedFolder();
    void slotRunSelectedRuleOnSelectedFolder();

    void slotRuleEnabledChecked();
    void slotDeleteRule();

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
    bool showRule( std::shared_ptr< Outlook::Rule > rule );
    bool editRule( std::shared_ptr< Outlook::Rule > rule );
    void updateActions();
    void reloadAll( bool andLoadServerInfo );

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
        bool enabled = !running();
        if ( !enabled )
        {
            reasons.clear();
            reasons.emplace_back( false, "Cannot execute while processing" );
        }
        for ( auto &&ii : reasons )
        {
            enabled = enabled && ii.first;
        }

        item->setEnabled( enabled );
        auto itemText = item->property( "text" );
        QString toolTip;
        QString separator;

        if ( !enabled )
        {
            TReasons disabledReasons;
            for ( auto &&ii : reasons )
            {
                if ( !ii.first )
                    disabledReasons.push_back( ii );
            }

            if ( disabledReasons.size() == 1 )
            {
                toolTip = disabledReasons.front().second;
                separator = " - ";
            }
            else
            {
                for ( auto &&ii : disabledReasons )
                {
                    toolTip += "<li style=\"white-space:nowrap\">" + ii.second + "<li>\n";
                }
                toolTip = "<ul>\n" + toolTip + "</ul>";
                separator = ":";
            }
        }
        if ( itemText.isValid() )
            toolTip = itemText.toString() + separator + toolTip;
        item->setToolTip( toolTip );
    }
    template< typename T >
    void setEnabled()
    {
        auto children = this->findChildren< T >();
        for ( auto &&child : children )
        {
            if ( dynamic_cast< QAction * >( child ) && dynamic_cast< QAction * >( child )->menu() )
                continue;
            if ( dynamic_cast< QAbstractButton * >( child ) && ( dynamic_cast< QAbstractButton * >( child ) == fCancelButton ) )
                continue;
            setEnabled( child );
        }
    }

    CStatusProgress *getProgressBar( const QString &label );
    void setupStatusBar();

    CStatusProgress *addStatusBar( QString label, CWidgetWithStatus *object );

    void clearViews();
    std::unique_ptr< Ui::CMainWindow > fImpl;
    std::map< QString, CStatusProgress * > fProgressBars;
    std::optional< int > fNumWaitCursors;
    QPushButton *fCancelButton{ nullptr };
};

#endif
