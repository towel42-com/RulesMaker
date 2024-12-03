#ifndef EMAILVIEW_H
#define EMAILVIEW_H

#include <QWidget>
#include <memory>
namespace Ui
{
    class CEmailView;
}

class QModelIndex;
class CGroupedEmailModel;

class CEmailView : public QWidget
{
    Q_OBJECT

public:
    explicit CEmailView( QWidget *parent = nullptr );

    void init();

    ~CEmailView();

    void clear();
    void reload( bool notifyOnFinished );

    QStringList getRulesForSelection() const;

    void setOnlyProcessUnread( bool value );
    bool onlyProcessUnread() const;

    void setProcessAllEmailWhenLessThan200Emails( bool value );
    bool processAllEmailWhenLessThan200Emails() const;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedGrouping();
    void sigRuleSelected();
    void sigSetStatus( int curr, int max );

protected Q_SLOTS:
    void slotSelectionChanged();
    void slotItemDoubleClicked( const QModelIndex &idx );

protected:
    CGroupedEmailModel *fGroupedModel;
    std::unique_ptr< Ui::CEmailView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CONTACTSVIEW_H
