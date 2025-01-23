#ifndef RULESMODEL_H
#define RULESMODEL_H

#include "OutlookAPI/OutlookObj.h"

#include <QString>
#include <QStandardItemModel>

#include <memory>
#include <optional>
#include <unordered_map>
#include <unordered_set>

namespace Outlook
{
    class Rules;
    class Rule;
    class RuleConditions;
    class Folder;

}

class CRulesModel : public QStandardItemModel
{
    Q_OBJECT;

public:
    explicit CRulesModel( QObject *parent );

    virtual ~CRulesModel();

    void reload();
    void clear();

    void update();

    bool ruleSelected( const QModelIndex &index ) const;
    bool ruleSelected( const QStandardItem *item ) const;

    QStandardItem *getRuleItem( const QModelIndex &index ) const;
    QStandardItem *getRuleItem( const QStandardItem *item ) const;

    COutlookObj< Outlook::Rule > getRule( const QModelIndex &index ) const;
    COutlookObj< Outlook::Rule > getRule( const QStandardItem *item ) const;

    QString summary() const;

    virtual bool hasChildren( const QModelIndex &parent ) const override;
    virtual void fetchMore( const QModelIndex &parent ) override;
    virtual bool canFetchMore( const QModelIndex &parent ) const override;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotLoadNextRule();
    void slotRuleAdded( const COutlookObj< Outlook::Rule > & rule );
    void slotRuleChanged( const COutlookObj< Outlook::Rule > & rule );
    void slotRuleDeleted( const COutlookObj< Outlook::Rule > & rule );

private:
    void updateAllRules();
    bool beenLoaded( const QModelIndex &parent ) const;

    void loadRules();

    bool loadRule( const COutlookObj< Outlook::Rule > & rule, QStandardItem *ruleItem = nullptr );
    bool updateRule( const COutlookObj< Outlook::Rule > & rule );

    std::pair< COutlookObj< Outlook::Rules >, int > fRules{ COutlookObj< Outlook::Rules >{}, 0 };
    std::unordered_map< QStandardItem *, COutlookObj< Outlook::Rule > > fRuleMap;
    std::unordered_map< COutlookObj< Outlook::Rule >, QStandardItem * > fReverseRuleMap;
    int fCurrPos{ 1 };
};

#endif
