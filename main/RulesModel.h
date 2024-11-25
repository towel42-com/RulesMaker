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

    class AccountRuleCondition;
    class AddressRuleCondition;
    class CategoryRuleCondition;
    class FormNameRuleCondition;
    class FromRssFeedRuleCondition;
    class ImportanceRuleCondition;
    class RuleCondition;
    class SenderInAddressListRuleCondition;
    class SensitivityRuleCondition;
    class TextRuleCondition;
    class ToOrFromRuleCondition;

    class AssignToCategoryRuleAction;
    class MarkAsTaskRuleAction;
    class MoveOrCopyRuleAction;
    class NewItemAlertRuleAction;
    class PlaySoundRuleAction;
    class RuleAction;
    class SendRuleAction;
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

    void runRule( const QModelIndex &index ) const;
    void runRule( QStandardItem *item ) const;

    std::shared_ptr< Outlook::Rule > getRule( const QModelIndex &index ) const;
    std::shared_ptr< Outlook::Rule > getRule( QStandardItem *item ) const;

    bool addRule( const QString &destFolder, const QStringList &rules, QStringList &msgs );
    bool addToRule( std::shared_ptr< Outlook::Rule > rule, const QStringList &rules, QStringList &msgs );

    virtual bool hasChildren( const QModelIndex &parent ) const override;
    virtual void fetchMore( const QModelIndex &parent ) override;
    virtual bool canFetchMore( const QModelIndex &parent ) const override;


Q_SIGNALS:
    void sigFinishedLoading();

private:
    bool beenLoaded( const QModelIndex &parent ) const;
    bool beenLoaded( QStandardItem *parent ) const;

    void loadRules();

    bool loadRule( std::shared_ptr< Outlook::Rule > rule );

    void loadRuleData( QStandardItem *ruleItem, std::shared_ptr< Outlook::Rule > rule );

    bool updateRule( std::shared_ptr< Outlook::Rule > rule );

    void addAttribute( QStandardItem *parent, const QString &label, const QString &value );
    void addAttribute( QStandardItem *parent, const QString &label, QStringList value, const QString &separator );
    void addAttribute( QStandardItem *parent, const QString &label, bool value );
    void addAttribute( QStandardItem *parent, const QString &label, int value );
    void addAttribute( QStandardItem *parent, const QString &label, const char *value );

    void addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule );
    void addExceptions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule );
    void addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule, bool exceptions );

    bool addCondition( QStandardItem *parent, Outlook::AccountRuleCondition *condition );
    bool addCondition( QStandardItem *parent, Outlook::RuleCondition *condition, const QString &ruleName );
    bool addCondition( QStandardItem *parent, Outlook::TextRuleCondition *condition, const QString &ruleName );
    bool addCondition( QStandardItem *parent, Outlook::CategoryRuleCondition *condition, const QString &ruleName );
    bool addCondition( QStandardItem *parent, Outlook::ToOrFromRuleCondition *condition, bool from );
    bool addCondition( QStandardItem *parent, Outlook::FormNameRuleCondition *condition );
    bool addCondition( QStandardItem *parent, Outlook::FromRssFeedRuleCondition *condition );
    bool addCondition( QStandardItem *parent, Outlook::ImportanceRuleCondition *condition );
    bool addCondition( QStandardItem *parent, Outlook::AddressRuleCondition *condition );
    bool addCondition( QStandardItem *parent, Outlook::SenderInAddressListRuleCondition *condition );
    bool addCondition( QStandardItem *parent, Outlook::SensitivityRuleCondition *condition );

    void addActions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule );
    bool addAction( QStandardItem *parent, Outlook::AssignToCategoryRuleAction *action );
    bool addAction( QStandardItem *parent, Outlook::MarkAsTaskRuleAction *action );
    bool addAction( QStandardItem *parent, Outlook::MoveOrCopyRuleAction *action, const QString &actionName );
    bool addAction( QStandardItem *parent, Outlook::NewItemAlertRuleAction *action );
    bool addAction( QStandardItem *parent, Outlook::PlaySoundRuleAction *action );
    bool addAction( QStandardItem *parent, Outlook::RuleAction *action, const QString &actionName );
    bool addAction( QStandardItem *parent, Outlook::SendRuleAction *action, const QString &actionName );

    std::shared_ptr< Outlook::Rules > fRules{ nullptr };
    std::unordered_map< QStandardItem *, std::shared_ptr< Outlook::Rule > > fRuleMap;
    std::unordered_set< QStandardItem * > fBeenLoaded;
};

#endif
