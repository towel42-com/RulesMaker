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
}

class QModelIndex;
class CRulesModel;

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
    void runSelectedRule() const;
    std::shared_ptr< Outlook::Rule > selectedRule() const;

    bool addRule( const QString &destFolder, const QStringList &rules, QStringList &msgs );
    bool addToSelectedRule( const QStringList &rules, QStringList &msgs );

Q_SIGNALS:
    void sigFinishedLoading();
    void sigRuleSelected();

protected Q_SLOTS:
    void slotItemSelected( const QModelIndex &index );

protected:
    CRulesModel *fModel{ nullptr };
    std::unique_ptr< Ui::CRulesView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CRulesView_H
