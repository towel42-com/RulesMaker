#include "OutlookAPI.h"
#include "OutlookAPI_pri.h"
#include "EmailAddress.h"

#include "MSOUTL.h"

#include <QRegularExpression>

std::list< QStringList > COutlookAPI::getConditionalStringList( const COutlookObj< Outlook::Rule > &rule, bool exceptions, EWrapperMode wrapperMode, bool includeSender )
{
    if ( !rule )
        return {};

    auto conditions = exceptions ? rule->Exceptions() : rule->Conditions();
    if ( !conditions )
        return {};

    std::list< QStringList > retVal;
    retVal.push_back( conditionNames( conditions->Account(), "Account", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->AnyCategory(), "AnyCategory", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->Body(), "Body", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->BodyOrSubject(), "BodyOrSubject", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->CC(), "CC", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->Category(), "Category", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->FormName(), "FormName", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->From(), "From", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->FromAnyRSSFeed(), "FromAnyRSSFeed", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->FromRssFeed(), "FromRssFeed", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->HasAttachment(), "HasAttachment", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->Importance(), "Importance", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->MeetingInviteOrUpdate(), "MeetingInviteOrUpdate", wrapperMode ) );
    retVal.push_back( conditionNamesForMsgHeader( conditions->MessageHeader(), "MessageHeader", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->NotTo(), "NotTo", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->OnLocalMachine(), "OnLocalMachine", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->OnOtherMachine(), "OnOtherMachine", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->OnlyToMe(), "OnlyToMe", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->RecipientAddress(), "RecipientAddress", wrapperMode ) );
    if ( includeSender )
        retVal.push_back( conditionNames( conditions->SenderAddress(), "SenderAddress", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->SenderInAddressList(), "SenderInAddressList", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->Sensitivity(), "Sensitivity", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->SentTo(), "SentTo", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->Subject(), "Subject", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->ToMe(), "ToMe", wrapperMode ) );
    retVal.push_back( conditionNames( conditions->ToOrCc(), "ToOrCc", wrapperMode ) );

    retVal.remove_if( []( const QStringList &list ) { return list.isEmpty(); } );
    return retVal;
}

QString COutlookAPI::rawRuleNameForRule( const COutlookObj< Outlook::Rule > &rule )
{
    QStringList addOns;
    if ( !rule )
        return {};
    return rule->Name();
}

std::optional< QString > COutlookAPI::getDestFolderNameForRule( const COutlookObj< Outlook::Rule > &rule, bool moveOnly )
{
    if ( !rule )
        return {};

    auto actions = rule->Actions();
    if ( !actions )
        return {};

    Outlook::MAPIFolder *destFolder = nullptr;
    auto mvToFolderAction = actions->MoveToFolder();
    auto cpToFolderAction = moveOnly ? actions->CopyToFolder() : nullptr;

    if ( mvToFolderAction && mvToFolderAction->Enabled() )
    {
        destFolder = mvToFolderAction->Folder();
    }
    else if ( cpToFolderAction && cpToFolderAction->Enabled() )
        destFolder = cpToFolderAction->Folder();

    if ( !destFolder )
        return {};

    return ruleNameForFolder( reinterpret_cast< Outlook::MAPIFolder * >( destFolder ) );
}

QString COutlookAPI::ruleNameForRule( const COutlookObj< Outlook::Rule > &rule, bool forDisplay )
{
    QStringList addOns;
    if ( !rule )
    {
        addOns << "INV-NULLPTR";
    }
    auto actions = rule ? rule->Actions() : nullptr;
    if ( !actions )
    {
        addOns << "INV-NOACTIONS";
    }

    QString ruleName;
    if ( forDisplay && rule )
        ruleName = getDisplayName( rule );
    else
    {
        auto destFolderName = getDestFolderNameForRule( rule, true );
        if ( destFolderName.has_value() )
            ruleName = destFolderName.value();
        else
            ruleName = "INV-NODESTFOLDER";
    }

    QString conditionals;
    QString exceptions;
    if ( !forDisplay )
    {
        auto join = []( const std::list< QStringList > &list ) -> QString
        {
            QString retVal;

            QStringList tmp;
            if ( list.size() == 1 )
            {
                tmp << list.front().join( " or " );
            }
            else
            {
                for ( auto &&ii : list )
                {
                    if ( ii.isEmpty() )
                        continue;
                    tmp << "(" + ii.join( " or " ) + ")";
                }
            }
            return tmp.join( " and " );
        };
        conditionals = join( getConditionalStringList( rule, false, EWrapperMode::eParenIndividual, false ) );
        exceptions = join( getConditionalStringList( rule, true, EWrapperMode::eParenIndividual, false ) );
    }

    if ( ruleName.isEmpty() )
        ruleName = "<UNNAMED RULE>";

    addOns.removeAll( QString() );
    addOns.sort();

    auto suffixes = QStringList() << ruleName << addOns.join( " " ) << conditionals << exceptions;
    suffixes.removeAll( QString() );
    for ( auto &&ii : suffixes )
        ii = ii.trimmed();

    return suffixes.join( " " ).trimmed();
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

QStringList conditionNames( Outlook::SensitivityRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toString( condition->Sensitivity() ), wrapperMode );
}

QStringList conditionNames( Outlook::SenderInAddressListRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto addresses = COutlookAPI::instance()->getEmailAddresses( condition->AddressList() );

    return conditionRuleNameBase( condition, conditionStr, toStringList( addresses ), wrapperMode );
}

QStringList conditionNames( Outlook::AddressRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toStringList( condition->Address() ), wrapperMode );
}

QStringList conditionNames( Outlook::ImportanceRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toString( condition->Importance() ), wrapperMode );
}

QStringList conditionNames( Outlook::FromRssFeedRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toStringList( condition->FromRssFeed() ), wrapperMode );
}

QStringList conditionNames( Outlook::FormNameRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toStringList( condition->FormName() ), wrapperMode );
}

QStringList conditionNames( Outlook::ToOrFromRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toStringList( COutlookAPI::getEmailAddresses( condition->Recipients() ) ), wrapperMode );
}

QStringList conditionNames( Outlook::CategoryRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toStringList( condition->Categories() ), wrapperMode );
}

QStringList conditionNames( Outlook::TextRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toStringList( condition->Text() ), wrapperMode );
}

QStringList conditionNamesForMsgHeader( Outlook::TextRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto text = toStringList( condition->Text() );

    QRegularExpression regex( R"__(From: "(.*)")__" );

    for ( auto &&ii : text )
    {
        auto match = regex.match( ii );
        if ( match.hasMatch() )
        {
            auto fromInfo = match.captured( 1 );
            ii = QString( R"(From: %1)" ).arg( match.captured( 1 ) );
        }
    }
    text.removeAll( QString() );
    text.removeDuplicates();

    return conditionRuleNameBase( condition, conditionStr, text, wrapperMode );
}
QStringList conditionNames( Outlook::RuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, "Yes", wrapperMode );
}

QStringList conditionNames( Outlook::AccountRuleCondition *condition, const QString &conditionStr, EWrapperMode wrapperMode )
{
    if ( !condition || !condition->Enabled() )
        return {};

    return conditionRuleNameBase( condition, conditionStr, toString( condition->ConditionType() ), wrapperMode );
}

QStringList COutlookAPI::getActionStrings( const COutlookObj< Outlook::Rule > &rule )
{
    if ( !rule )
        return {};

    if ( !rule )
        return {};

    auto actions = rule->Actions();
    if ( !actions )
        return {};

    QStringList retVal;
    retVal << actionName( actions->AssignToCategory() );
    retVal << actionName( actions->MarkAsTask() );
    retVal << actionName( actions->CopyToFolder(), "Copy to Folder" );
    retVal << actionName( actions->MoveToFolder(), "Move to Folder" );
    retVal << actionName( actions->NewItemAlert() );
    retVal << actionName( actions->PlaySound() );
    retVal << actionName( actions->ClearCategories(), "Clear Categories" );
    retVal << actionName( actions->Delete(), "Delete" );
    retVal << actionName( actions->DeletePermanently(), "Delete Permanently" );
    retVal << actionName( actions->DesktopAlert(), "Desktop Alert" );
    retVal << actionName( actions->NotifyDelivery(), "Notify Delivery" );
    retVal << actionName( actions->NotifyRead(), "Notify Read" );
    retVal << actionName( actions->Stop(), "Stop" );
    retVal << actionName( actions->CC(), "Send as CC" );
    retVal << actionName( actions->Forward(), "Forward" );
    retVal << actionName( actions->ForwardAsAttachment(), "Forward as Attachment" );
    retVal << actionName( actions->Redirect(), "Redirect" );

    retVal.removeAll( QString() );
    retVal.sort();

    return retVal;
}

QString actionName( Outlook::AssignToCategoryRuleAction *action )
{
    if ( !action )
        return {};
    if ( !action->Enabled() )
        return false;

    return QString( "Set Categories To: %1" ).arg( toString( action->Categories(), " and " ) );
}

QString actionName( Outlook::MarkAsTaskRuleAction *action )
{
    if ( !action )
        return {};
    if ( !action->Enabled() )
        return false;
    return QString( "Mark as Task: Yes - %1" ).arg( toString( action->MarkInterval() ) );
}

QString actionName( Outlook::MoveOrCopyRuleAction *action, const QString &actionName )
{
    if ( !action )
        return {};
    if ( !action->Enabled() )
        return false;
    return QString( "%1: %2" ).arg( actionName ).arg( action->Folder()->FullFolderPath() );
}

QString actionName( Outlook::NewItemAlertRuleAction *action )
{
    if ( !action )
        return {};
    if ( !action->Enabled() )
        return false;
    return QString( "New Item Alert: %1" ).arg( action->Text() );
}

QString actionName( Outlook::PlaySoundRuleAction *action )
{
    if ( !action )
        return {};
    if ( !action->Enabled() )
        return false;
    return QString( "Play Sound: \"%1\"" ).arg( action->FilePath() );
}

QString actionName( Outlook::RuleAction *action, const QString &actionName )
{
    if ( !action )
        return {};
    if ( !action->Enabled() )
        return false;
    return QString( "%1: Yes" ).arg( actionName );
}

QString actionName( Outlook::SendRuleAction *action, const QString &actionName )
{
    if ( !action )
        return {};
    if ( !action->Enabled() )
        return false;
    auto recipients = COutlookAPI::getEmailAddresses( action->Recipients() );

    return QString( "%1: %2" ).arg( actionName ).arg( toStringList( recipients ).join( " and " ) );
}
