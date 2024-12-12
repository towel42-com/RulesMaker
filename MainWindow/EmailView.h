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

    QStringList getRulesForSelection() const;

    QString getSelectedDisplayName() const;

    
Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedGrouping();
    void sigRuleSelected();

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
