#ifndef RULESVIEW_H
#define RULESVIEW_H

#include "WidgetWithStatus.h"
#include <memory>
namespace Ui
{
    class CRulesView;
}

namespace Outlook
{
    class Rule;
    class Folder;
}

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

Q_SIGNALS:
    void sigFinishedLoading();
    void sigRuleSelected();

protected Q_SLOTS:
    void slotItemSelected( const QModelIndex &index );
    void slotDeleteCurrent();

protected:
    QModelIndex currentIndex() const;
    QModelIndex sourceIndex( const QModelIndex &idx ) const;

    CRulesModel *fModel{ nullptr };
    CListFilterModel *fFilterModel{ nullptr };
    std::unique_ptr< Ui::CRulesView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CRulesView_H