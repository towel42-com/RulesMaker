#ifndef OUTLOOKAPI_PRI_H
#define OUTLOOKAPI_PRI_H

#include <QString>
#include <memory>
#include <optional>
#include <type_traits>

class QStandardItem;
enum class EWrapperMode;
namespace Outlook
{
    class Application;
    class _Application;
    class NameSpace;
    class _NameSpace;
    class Account;
    class _Account;
    class Folder;
    class MAPIFolder;
    class MailItem;
    class _MailItem;
    class AddressEntry;
    class AddressEntries;
    class Recipient;
    class Rules;
    class _Rules;
    class RuleConditions;
    class RuleActions;
    class Rule;
    class _Rule;
    class Recipients;
    class AddressList;
    class Items;
    class _Items;

    enum class OlImportance;
    enum class OlItemType;
    enum class OlMailRecipientType;
    enum class OlObjectClass;
    enum class OlRuleConditionType;
    enum class OlSensitivity;
    enum class OlMarkInterval;
    enum class OlDefaultFolders;
    enum class OlAddressEntryUserType;
    enum class OlDisplayType;

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

QString actionName( Outlook::AssignToCategoryRuleAction *action );
QString actionName( Outlook::MarkAsTaskRuleAction *action );
QString actionName( Outlook::MoveOrCopyRuleAction *action, const QString &actionName );
QString actionName( Outlook::NewItemAlertRuleAction *action );
QString actionName( Outlook::PlaySoundRuleAction *action );
QString actionName( Outlook::RuleAction *action, const QString &actionName );
QString actionName( Outlook::SendRuleAction *action, const QString &actionName );

QStringList conditionNames( Outlook::AccountRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::AddressRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::CategoryRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::FormNameRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::FromRssFeedRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::ImportanceRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::RuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::SenderInAddressListRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::SensitivityRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::TextRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNames( Outlook::ToOrFromRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );
QStringList conditionNamesForMsgHeader( Outlook::TextRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode );

bool actionEqual( Outlook::AssignToCategoryRuleAction *lhsAction, Outlook::AssignToCategoryRuleAction *rhsAction );
bool actionEqual( Outlook::MarkAsTaskRuleAction *lhsAction, Outlook::MarkAsTaskRuleAction *rhsAction );
bool actionEqual( Outlook::MoveOrCopyRuleAction *lhsAction, Outlook::MoveOrCopyRuleAction *rhsAction );
bool actionEqual( Outlook::NewItemAlertRuleAction *lhsAction, Outlook::NewItemAlertRuleAction *rhsAction );
bool actionEqual( Outlook::PlaySoundRuleAction *lhsAction, Outlook::PlaySoundRuleAction *rhsAction );
bool actionEqual( Outlook::RuleAction *lhsAction, Outlook::RuleAction *rhsAction );
bool actionEqual( Outlook::SendRuleAction *lhsAction, Outlook::SendRuleAction *rhsAction );
bool actionsEqual( Outlook::RuleActions *lhsAction, Outlook::RuleActions *rhsAction );

bool loadAction( QStandardItem *parent, Outlook::AssignToCategoryRuleAction *action );
bool loadAction( QStandardItem *parent, Outlook::MarkAsTaskRuleAction *action );
bool loadAction( QStandardItem *parent, Outlook::MoveOrCopyRuleAction *action, const QString &actionName );
bool loadAction( QStandardItem *parent, Outlook::NewItemAlertRuleAction *action );
bool loadAction( QStandardItem *parent, Outlook::PlaySoundRuleAction *action );
bool loadAction( QStandardItem *parent, Outlook::RuleAction *action, const QString &actionName );
bool loadAction( QStandardItem *parent, Outlook::SendRuleAction *action, const QString &actionName );
void loadActions( QStandardItem *parent, const COutlookObj< Outlook::Rule > & rule );

bool loadCondition( QStandardItem *parent, Outlook::AccountRuleCondition *condition );
bool loadCondition( QStandardItem *parent, Outlook::AddressRuleCondition *condition );
bool loadCondition( QStandardItem *parent, Outlook::CategoryRuleCondition *condition, const QString &ruleName );
bool loadCondition( QStandardItem *parent, Outlook::FormNameRuleCondition *condition );
bool loadCondition( QStandardItem *parent, Outlook::FromRssFeedRuleCondition *condition );
bool loadCondition( QStandardItem *parent, Outlook::ImportanceRuleCondition *condition );
bool loadCondition( QStandardItem *parent, Outlook::RuleCondition *condition, const QString &ruleName );
bool loadCondition( QStandardItem *parent, Outlook::SenderInAddressListRuleCondition *condition );
bool loadCondition( QStandardItem *parent, Outlook::SensitivityRuleCondition *condition );
bool loadCondition( QStandardItem *parent, Outlook::TextRuleCondition *condition, const QString &ruleName );
bool loadCondition( QStandardItem *parent, Outlook::ToOrFromRuleCondition *condition, bool from );   // from or sentTo

void loadConditions( QStandardItem *parent, const COutlookObj< Outlook::Rule > &rule );
void loadConditions( QStandardItem *parent, const COutlookObj< Outlook::Rule > &rule, bool exceptions );
void loadExceptions( QStandardItem *parent, const COutlookObj< Outlook::Rule > &rule );

bool conditionEqual( Outlook::AccountRuleCondition *lhsCondition, Outlook::AccountRuleCondition *rhsCondition );
bool conditionEqual( Outlook::AddressRuleCondition *lhsCondition, Outlook::AddressRuleCondition *rhsCondition );
bool conditionEqual( Outlook::CategoryRuleCondition *lhsCondition, Outlook::CategoryRuleCondition *rhsCondition );
bool conditionEqual( Outlook::FormNameRuleCondition *lhsCondition, Outlook::FormNameRuleCondition *rhsCondition );
bool conditionEqual( Outlook::FromRssFeedRuleCondition *lhsCondition, Outlook::FromRssFeedRuleCondition *rhsCondition );
bool conditionEqual( Outlook::ImportanceRuleCondition *lhsCondition, Outlook::ImportanceRuleCondition *rhsCondition );
bool conditionEqual( Outlook::RuleCondition *lhsCondition, Outlook::RuleCondition *rhsCondition );
bool conditionEqual( Outlook::SenderInAddressListRuleCondition *lhsCondition, Outlook::SenderInAddressListRuleCondition *rhsCondition );
bool conditionEqual( Outlook::SensitivityRuleCondition *lhsCondition, Outlook::SensitivityRuleCondition *rhsCondition );
bool conditionEqual( Outlook::TextRuleCondition *lhsCondition, Outlook::TextRuleCondition *rhsCondition );
bool conditionEqual( Outlook::ToOrFromRuleCondition *lhsCondition, Outlook::ToOrFromRuleCondition *rhsCondition );
std::optional< int > numConditionsDifferent( Outlook::RuleConditions *lhs, Outlook::RuleConditions *rhs );

void loadAttribute( QStandardItem *parent, const QString &label, const QStringList &value, const QString &separator );
void loadAttribute( QStandardItem *parent, const QString &label, const TEmailAddressList &value, const QString &separator );
void loadAttribute( QStandardItem *parent, const QString &label, bool value );
void loadAttribute( QStandardItem *parent, const QString &label, const QString &value );
void loadAttribute( QStandardItem *parent, const QString &label, const char *value );
void loadAttribute( QStandardItem *parent, const QString &label, int value );

void copyAction( Outlook::AssignToCategoryRuleAction *retVal, Outlook::AssignToCategoryRuleAction *sourceAction );
void copyAction( Outlook::MarkAsTaskRuleAction *retVal, Outlook::MarkAsTaskRuleAction *sourceAction );
void copyAction( Outlook::MoveOrCopyRuleAction *retVal, Outlook::MoveOrCopyRuleAction *sourceAction );
void copyAction( Outlook::NewItemAlertRuleAction *retVal, Outlook::NewItemAlertRuleAction *sourceAction );
void copyAction( Outlook::PlaySoundRuleAction *retVal, Outlook::PlaySoundRuleAction *sourceAction );
void copyAction( Outlook::RuleAction *retVal, Outlook::RuleAction *sourceAction );
void copyAction( Outlook::SendRuleAction *retVal, Outlook::SendRuleAction *sourceAction );
void copyActions( COutlookObj< Outlook::Rule > & retValRule, const COutlookObj< Outlook::Rule > &source );

void copyCondition( Outlook::AccountRuleCondition *retVal, Outlook::AccountRuleCondition *sourceCondition );
void copyCondition( Outlook::AddressRuleCondition *retVal, Outlook::AddressRuleCondition *sourceCondition );
void copyCondition( Outlook::CategoryRuleCondition *retVal, Outlook::CategoryRuleCondition *sourceCondition );
void copyCondition( Outlook::FormNameRuleCondition *retVal, Outlook::FormNameRuleCondition *sourceCondition );
void copyCondition( Outlook::FromRssFeedRuleCondition *retVal, Outlook::FromRssFeedRuleCondition *sourceCondition );
void copyCondition( Outlook::ImportanceRuleCondition *retVal, Outlook::ImportanceRuleCondition *sourceCondition );
void copyCondition( Outlook::RuleCondition *retVal, Outlook::RuleCondition *sourceCondition );
void copyCondition( Outlook::SenderInAddressListRuleCondition *retVal, Outlook::SenderInAddressListRuleCondition *sourceCondition );
void copyCondition( Outlook::SensitivityRuleCondition *retVal, Outlook::SensitivityRuleCondition *sourceCondition );
void copyCondition( Outlook::TextRuleCondition *retVal, Outlook::TextRuleCondition *sourceCondition );
void copyCondition( Outlook::ToOrFromRuleCondition *retVal, Outlook::ToOrFromRuleCondition *sourceCondition );
void copyConditions( COutlookObj< Outlook::Rule > &retValRule, const COutlookObj< Outlook::Rule > &source, bool exceptions );

void mergeCondition( Outlook::AddressRuleCondition *lhsCondition, Outlook::AddressRuleCondition *rhsCondition );
void mergeCondition( Outlook::CategoryRuleCondition *lhsCondition, Outlook::CategoryRuleCondition *rhsCondition );
void mergeCondition( Outlook::FormNameRuleCondition *lhsCondition, Outlook::FormNameRuleCondition *rhsCondition );
void mergeCondition( Outlook::FromRssFeedRuleCondition *lhsCondition, Outlook::FromRssFeedRuleCondition *rhsCondition );
void mergeCondition( Outlook::TextRuleCondition *lhsCondition, Outlook::TextRuleCondition *rhsCondition );
void mergeCondition( Outlook::ToOrFromRuleCondition *lhsCondition, Outlook::ToOrFromRuleCondition *rhsCondition );
void mergeConditions( Outlook::RuleConditions *lhs, Outlook::RuleConditions *rhs );

#endif
