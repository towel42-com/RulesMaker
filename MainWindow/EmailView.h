#ifndef EMAILVIEW_H
#define EMAILVIEW_H

#include "WidgetWithStatus.h"
#include <memory>
#include <QModelIndexList>
namespace Ui
{
    class CEmailView;
}

class CEmailModel;

class CEmailView : public CWidgetWithStatus
{
    Q_OBJECT

public:
    explicit CEmailView( QWidget *parent = nullptr );
    ~CEmailView();

    void init();

    void clear();
    void clearSelection();
    void reload( bool notifyOnFinished );

    QStringList getMatchTextForSelection() const;
    QString getDisplayTextForSelection() const;

    QString getEmailDisplayNameForSelection() const;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedGrouping();
    void sigEmailSelected();

public Q_SLOTS:
    void slotRunningStateChanged( bool running );

protected Q_SLOTS:
    void slotSelectionChanged();
    void slotItemDoubleClicked( const QModelIndex &idx );

protected:
    QModelIndexList getSelectedRows() const;

    CEmailModel *fGroupedModel{ nullptr };
    std::unique_ptr< Ui::CEmailView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CONTACTSVIEW_H
