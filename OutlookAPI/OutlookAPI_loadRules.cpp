#include "OutlookAPI.h"
#include "OutlookAPI_pri.h"

#include "EmailAddress.h" 
#include "MSOUTL.h"

#include <QStandardItem>
#include <QRegularExpression>

void COutlookAPI::loadRuleData( QStandardItem *ruleItem, COutlookObj< Outlook::_Rule > rule, bool force )
{
    if ( ruleBeenLoaded( rule ) )
    {
        if ( !force )
            return;
        else
            ruleItem->removeRows( 0, ruleItem->rowCount() );
    }

    loadAttribute( ruleItem, "Name", rule->Name() );
    loadAttribute( ruleItem, "Enabled", rule->Enabled() );
    loadAttribute( ruleItem, "Execution Order", rule->ExecutionOrder() );
    loadAttribute( ruleItem, "Is Local", rule->IsLocalRule() );
    loadAttribute( ruleItem, "Rule Type", toString( rule->RuleType() ) );

    loadConditions( ruleItem, rule );
    loadExceptions( ruleItem, rule );
    loadActions( ruleItem, rule );
    fRuleBeenLoaded.insert( rule );
}

template< typename T >
static QStringList conditionRuleNameBase( T *condition, const QString &conditionStr, const QStringList &values, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    QStringList conditions;
    if ( wrapperMode == EWrapperMode::eAngleAll || wrapperMode == EWrapperMode::eParenAll )
    {
        conditions << conditionStr + "=" + values.join( " or " );
    }
    else
    {
        for ( auto &&ii : values )
        {
            conditions << conditionStr + "=" + ii;
        }
    }

    for ( auto &&ii : conditions )
    {
        switch ( wrapperMode )
        {
            case EWrapperMode::eAngleIndividual:
            case EWrapperMode::eAngleAll:
                ii = "<" + ii + ">";
                break;
            case EWrapperMode::eParenAll:
            case EWrapperMode::eParenIndividual:
                ii = "(" + ii + ")";
                break;
            default:
                break;
        }
    }
    return conditions;
}

template< typename T >
static QStringList conditionRuleNameBase( T *condition, const QString &conditionStr, const QString &value, EWrapperMode wrapperMode )
{
    return conditionRuleNameBase( condition, conditionStr, QStringList() << value, wrapperMode );
}

void loadConditions( QStandardItem *parent, const COutlookObj< Outlook::_Rule > & rule )
{
    return loadConditions( parent, rule, false );
}

void loadExceptions( QStandardItem *parent, const COutlookObj< Outlook::_Rule > & rule )
{
    return loadConditions( parent, rule, true );
}

void loadConditions( QStandardItem *parent, const COutlookObj< Outlook::_Rule > & rule, bool exceptions )
{
    if ( !rule )
        return;

    auto conditions = exceptions ? rule->Exceptions() : rule->Conditions();
    if ( !conditions )
        return;

    auto count = conditions->Count();
    if ( !count )
        return;
    auto folder = new QStandardItem( exceptions ? "Exceptions" : "Conditions" );
    auto found = false;

    found = loadCondition( folder, conditions->Account() ) || found;
    found = loadCondition( folder, conditions->AnyCategory(), "Any Category" ) || found;
    found = loadCondition( folder, conditions->Body(), "Body" ) || found;
    found = loadCondition( folder, conditions->BodyOrSubject(), "Body or Subject" ) || found;
    found = loadCondition( folder, conditions->CC(), "CC" ) || found;
    found = loadCondition( folder, conditions->Category(), "Category" ) || found;
    found = loadCondition( folder, conditions->FormName() ) || found;
    found = loadCondition( folder, conditions->From(), true ) || found;
    found = loadCondition( folder, conditions->FromAnyRSSFeed(), "From Any RSS Feed" ) || found;
    found = loadCondition( folder, conditions->FromRssFeed() ) || found;
    found = loadCondition( folder, conditions->HasAttachment(), "Has Attachment" ) || found;
    found = loadCondition( folder, conditions->Importance() ) || found;
    found = loadCondition( folder, conditions->MeetingInviteOrUpdate(), "Meeting Invite Or Update" ) || found;
    found = loadCondition( folder, conditions->MessageHeader(), "Message Header" ) || found;
    found = loadCondition( folder, conditions->NotTo(), "Not To" ) || found;
    found = loadCondition( folder, conditions->OnLocalMachine(), "On Local Machine" ) || found;
    found = loadCondition( folder, conditions->OnOtherMachine(), "On Other Machine" ) || found;
    found = loadCondition( folder, conditions->OnlyToMe(), "Only to Me" ) || found;
    found = loadCondition( folder, conditions->RecipientAddress() ) || found;
    found = loadCondition( folder, conditions->SenderAddress() ) || found;
    found = loadCondition( folder, conditions->SenderInAddressList() ) || found;
    found = loadCondition( folder, conditions->Sensitivity() ) || found;
    found = loadCondition( folder, conditions->SentTo(), "Sent To" ) || found;
    found = loadCondition( folder, conditions->Subject(), "Subject" ) || found;
    found = loadCondition( folder, conditions->ToMe(), "To Me" ) || found;
    found = loadCondition( folder, conditions->ToOrCc(), "To or CC" ) || found;

    if ( found )
        parent->appendRow( folder );
    else
        delete folder;
}

bool loadCondition( QStandardItem *parent, Outlook::AccountRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, "Condition Type", toString( condition->ConditionType() ) );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::RuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, ruleName, "Yes" );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::ToOrFromRuleCondition *condition, bool from )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    auto recipients = COutlookAPI::getEmailAddresses( condition->Recipients() );
    loadAttribute( parent, ( from ? "From" : "To" ), recipients, " or " );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::TextRuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, ruleName, toString( condition->Text(), " or " ) );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::CategoryRuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, ruleName, toString( condition->Categories(), " or " ) );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::FormNameRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, "Form Name", toString( condition->FormName(), " or " ) );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::FromRssFeedRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, "From RSS Feed", toString( condition->FromRssFeed(), " or " ) );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::ImportanceRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, "Importance", toString( condition->Importance() ) );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::AddressRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, "Address", toString( condition->Address(), " or " ) );
    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::SenderInAddressListRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    auto addresses = COutlookAPI::instance()->getEmailAddresses( condition->AddressList() );
    loadAttribute( parent, "Sender in Address List", addresses, " or " );

    return true;
}

bool loadCondition( QStandardItem *parent, Outlook::SensitivityRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    loadAttribute( parent, "Sensitivity", toString( condition->Sensitivity() ) );
    return true;
}

void loadActions( QStandardItem *parent, const COutlookObj< Outlook::_Rule > & rule )
{
    if ( !rule )
        return;

    auto actions = rule->Actions();
    if ( !actions )
        return;

    auto count = actions->Count();
    if ( !count )
        return;
    auto folder = new QStandardItem( "Actions" );
    auto found = false;

    found = loadAction( folder, actions->AssignToCategory() ) || found;
    found = loadAction( folder, actions->MarkAsTask() ) || found;
    found = loadAction( folder, actions->CopyToFolder(), "Copy to Folder" ) || found;
    found = loadAction( folder, actions->MoveToFolder(), "Move to Folder" ) || found;
    found = loadAction( folder, actions->NewItemAlert() ) || found;
    found = loadAction( folder, actions->PlaySound() ) || found;
    found = loadAction( folder, actions->ClearCategories(), "Clear Categories" ) || found;
    found = loadAction( folder, actions->Delete(), "Delete" ) || found;
    found = loadAction( folder, actions->DeletePermanently(), "Delete Permanently" ) || found;
    found = loadAction( folder, actions->DesktopAlert(), "Desktop Alert" ) || found;
    found = loadAction( folder, actions->NotifyDelivery(), "Notify Delivery" ) || found;
    found = loadAction( folder, actions->NotifyRead(), "Notify Read" ) || found;
    found = loadAction( folder, actions->Stop(), "Stop" ) || found;
    found = loadAction( folder, actions->CC(), "Send as CC" ) || found;
    found = loadAction( folder, actions->Forward(), "Forward" ) || found;
    found = loadAction( folder, actions->ForwardAsAttachment(), "Forward as Attachment" ) || found;
    found = loadAction( folder, actions->Redirect(), "Redirect" ) || found;

    if ( found )
        parent->appendRow( folder );
    else
        delete folder;
}

bool loadAction( QStandardItem *parent, Outlook::AssignToCategoryRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    loadAttribute( parent, "Set Categories To", toString( action->Categories(), " and " ) );
    return true;
}

bool loadAction( QStandardItem *parent, Outlook::MarkAsTaskRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    loadAttribute( parent, "Mark as Task:", QString( "%1 - %2" ).arg( action->FlagTo(), toString( action->MarkInterval() ) ) );
    return true;
}

bool loadAction( QStandardItem *parent, Outlook::MoveOrCopyRuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    loadAttribute( parent, actionName, action->Folder()->FullFolderPath() );
    return true;
}

bool loadAction( QStandardItem *parent, Outlook::NewItemAlertRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    loadAttribute( parent, "New Item Alert", action->Text() );
    return true;
}

bool loadAction( QStandardItem *parent, Outlook::PlaySoundRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;
    loadAttribute( parent, "Play Sound", '"' + action->FilePath() + '"' );
    return true;
}

bool loadAction( QStandardItem *parent, Outlook::RuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    loadAttribute( parent, actionName, "Yes" );
    return true;
}

bool loadAction( QStandardItem *parent, Outlook::SendRuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    auto recipients = COutlookAPI::getEmailAddresses( action->Recipients() );

    loadAttribute( parent, actionName, recipients, " and " );
    return true;
}

void loadAttribute( QStandardItem *parent, const QString &label, bool value )
{
    return loadAttribute( parent, label, value ? "Yes" : "No" );
}

void loadAttribute( QStandardItem *parent, const QString &label, int value )
{
    return loadAttribute( parent, label, QString::number( value ) );
}

void loadAttribute( QStandardItem *parent, const QString &label, const char *value )
{
    return loadAttribute( parent, label, QString( value ) );
}

void loadAttribute( QStandardItem *parent, const QString &label, const QStringList & value, const QString &separator )
{
    QStringList tmp;
    if ( value.size() > 1 )
    {
        for ( auto &&ii : value )
            tmp << '"' + ii + '"';
    }
    else
        tmp = value;
    auto text = tmp.join( separator );
    return loadAttribute( parent, label, text );
}

void loadAttribute( QStandardItem *parent, const QString &label, const TEmailAddressList & value, const QString &separator )
{
    QStringList tmp;
    if ( value.size() > 1 )
    {
        for ( auto &&ii : value )
            tmp << '"' + ii->toString() + '"';
    }
    else if ( !value.empty() )
        tmp << value.front()->toString();

    auto text = tmp.join( separator );
    return loadAttribute( parent, label, text );
}

void loadAttribute( QStandardItem *parent, const QString &label, const QString &value )
{
    auto keyItem = new QStandardItem( label + ":" );
    auto valueItem = new QStandardItem( value );
    parent->appendRow( { keyItem, valueItem } );
}
