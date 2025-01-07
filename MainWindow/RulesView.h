#ifndef RULESVIEW_H
#define RULESVIEW_H

#include "WidgetWithStatus.h"
#include <memory>
#include <optional>

namespace Ui
{
    class CRulesView;
}

namespace Outlook
{
    class Rule;
    class Folder;
}

enum class EFilterType;
class QModelIndex;
class CRulesModel;
class CListFilterModel;

class CRulesView : public CWidgetWithStatus
{
    Q_OBJECT

public:
    explicit CRulesView( QWidget *parent = nullptr );
    ~CRulesView();

    void init();

    void reload( bool notifyOnFinished );
    void clear();
    void clearSelection();

    bool ruleSelected() const;
    QString folderForSelectedRule() const;
    std::shared_ptr< Outlook::Rule > selectedRule() const;

    EFilterType filterTypeForSelectedRule() const;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigRuleSelected();

public Q_SLOTS:
    void slotRunningStateChanged( bool running );

protected Q_SLOTS:
    void slotItemSelected();

    void slotDeleteCurrent();
    void slotEnableCurrent();
    void slotDisableCurrent();
    void slotOptionsChanged();
    void slotRuleDoubleClicked();

protected:
    void updateButtons( const QModelIndex &index );
    void updateButtons( const std::shared_ptr< Outlook::Rule > &rule );
    QModelIndex selectedIndex() const;
    QModelIndex sourceIndex( const QModelIndex &idx ) const;

    CRulesModel *fModel{ nullptr };
    CListFilterModel *fFilterModel{ nullptr };
    std::unique_ptr< Ui::CRulesView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CRulesView_H
