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
    void reload();

    QStringList currentRule() const;

    void setOnlyGroupUnread( bool value );
    bool onlyGroupUnread() const;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedGrouping();
    void sigRuleSelected();

protected slots:
    void slotItemSelected( const QModelIndex &index );
    void slotItemDoubleClicked( const QModelIndex &idx );

protected:
    CGroupedEmailModel *fGroupedModel;
    std::unique_ptr< Ui::CEmailView > fImpl;
};

#endif   // CONTACTSVIEW_H
