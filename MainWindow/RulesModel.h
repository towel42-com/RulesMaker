#ifndef RULESMODEL_H
#define RULESMODEL_H

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

    QStandardItem *getRuleItem( const QModelIndex &index ) const;
    QStandardItem *getRuleItem( QStandardItem *item ) const;

    std::shared_ptr< Outlook::Rule > getRule( const QModelIndex &index ) const;
    std::shared_ptr< Outlook::Rule > getRule( QStandardItem *item ) const;

    virtual bool hasChildren( const QModelIndex &parent ) const override;
    virtual void fetchMore( const QModelIndex &parent ) override;
    virtual bool canFetchMore( const QModelIndex &parent ) const override;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigSetStatus( int curr, int max );

private Q_SLOTS:
    void slotLoadNextRule();
    void slotRuleAdded( std::shared_ptr< Outlook::Rule > rule );
    void slotRuleChanged( std::shared_ptr< Outlook::Rule > rule );
    void slotRuleDeleted( std::shared_ptr< Outlook::Rule > rule );

private:
    void updateAllRules();
    bool beenLoaded( const QModelIndex &parent ) const;

    void loadRules();

    bool loadRule( std::shared_ptr< Outlook::Rule > rule, QStandardItem * ruleItem = nullptr );
    bool updateRule( std::shared_ptr< Outlook::Rule > rule );

    std::pair< std::shared_ptr< Outlook::Rules >, int > fRules{ nullptr, 0 };
    std::unordered_map< QStandardItem *, std::shared_ptr< Outlook::Rule > > fRuleMap;
    std::unordered_map< std::shared_ptr< Outlook::Rule >, QStandardItem * > fReverseRuleMap;
    int fCurrPos{ 1 };
};

#endif
