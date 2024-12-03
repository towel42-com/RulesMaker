#ifndef RULESVIEW_H
#define RULESVIEW_H

#include <QWidget>
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

class CRulesView : public QWidget
{
    Q_OBJECT

public:
    explicit CRulesView( QWidget *parent = nullptr );

    void init();

    ~CRulesView();

    void reload( bool notifyOnFinished );
    void clear();

    bool ruleSelected() const;;
    void runSelectedRule() const;
    std::shared_ptr< Outlook::Rule > selectedRule() const;

    bool addRule( const QString &destFolder, const QStringList &rules, QStringList &msgs );
    bool addToSelectedRule( const QStringList &rules, QStringList &msgs );

Q_SIGNALS:
    void sigFinishedLoading();
    void sigRuleSelected();
    void sigSetStatus( int curr, int max );

protected Q_SLOTS:
    void slotItemSelected( const QModelIndex &index );

protected:
    std::shared_ptr< CRulesModel > fModel;
    std::unique_ptr< Ui::CRulesView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CRulesView_H
