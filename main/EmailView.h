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
    ~CEmailView();

    void init();

    void clear();
    void clearSelection();
    void reload( bool notifyOnFinished );

    QStringList getRulesForSelection() const;
    
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
