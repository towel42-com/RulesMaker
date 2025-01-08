#include "OutlookAPI.h"
#include "OutlookAPI_pri.h"

#include "MSOUTL.h"

bool actionEqual( Outlook::AssignToCategoryRuleAction *lhsAction, Outlook::AssignToCategoryRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    if ( lhsAction->Enabled() != rhsAction->Enabled() )
        return false;

    if ( !lhsAction->Enabled() || !lhsAction->Enabled() )
        return true;

    if ( !equal( toStringList( lhsAction->Categories() ), toStringList( rhsAction->Categories() ) ) )
        return false;
    return true;
}

bool actionEqual( Outlook::MarkAsTaskRuleAction *lhsAction, Outlook::MarkAsTaskRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    if ( lhsAction->Enabled() != rhsAction->Enabled() )
        return false;

    if ( !lhsAction->Enabled() || !lhsAction->Enabled() )
        return true;

    if ( lhsAction->FlagTo() != rhsAction->FlagTo() )
        return false;

    if ( lhsAction->MarkInterval() != rhsAction->MarkInterval() )
        return false;
    return true;
}

bool actionEqual( Outlook::MoveOrCopyRuleAction *lhsAction, Outlook::MoveOrCopyRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    if ( lhsAction->Enabled() != rhsAction->Enabled() )
        return false;

    if ( !lhsAction->Enabled() || !lhsAction->Enabled() )
        return true;

    if ( lhsAction->Folder()->FullFolderPath() != rhsAction->Folder()->FullFolderPath() )
        return false;

    return true;
}

bool actionEqual( Outlook::NewItemAlertRuleAction *lhsAction, Outlook::NewItemAlertRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    if ( lhsAction->Enabled() != rhsAction->Enabled() )
        return false;

    if ( !lhsAction->Enabled() || !lhsAction->Enabled() )
        return true;

    if ( lhsAction->Text() != rhsAction->Text() )
        return false;

    return true;
}

bool actionEqual( Outlook::PlaySoundRuleAction *lhsAction, Outlook::PlaySoundRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    if ( lhsAction->Enabled() != rhsAction->Enabled() )
        return false;

    if ( !lhsAction->Enabled() || !lhsAction->Enabled() )
        return true;

    if ( lhsAction->FilePath() != rhsAction->FilePath() )
        return false;

    return true;
}

bool actionEqual( Outlook::RuleAction *lhsAction, Outlook::RuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    if ( lhsAction->Enabled() != rhsAction->Enabled() )
        return false;

    if ( !lhsAction->Enabled() || !lhsAction->Enabled() )
        return true;

    return true;
}

bool actionEqual( Outlook::SendRuleAction *lhsAction, Outlook::SendRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    if ( lhsAction->Enabled() != rhsAction->Enabled() )
        return false;

    if ( !lhsAction->Enabled() || !lhsAction->Enabled() )
        return true;

    auto lhsRecipients = COutlookAPI::getEmailAddresses( lhsAction->Recipients(), {}, false );
    auto rhsRecipients = COutlookAPI::getEmailAddresses( rhsAction->Recipients(), {}, false );
    if ( !equal( lhsRecipients, rhsRecipients ) )
        return false;

    return true;
}

bool actionsEqual( Outlook::RuleActions *lhsAction, Outlook::RuleActions *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return false;

    auto retVal = true;

    retVal = retVal && actionEqual( lhsAction->AssignToCategory(), rhsAction->AssignToCategory() );
    retVal = retVal && actionEqual( lhsAction->MarkAsTask(), rhsAction->MarkAsTask() );
    retVal = retVal && actionEqual( lhsAction->CopyToFolder(), rhsAction->CopyToFolder() );
    retVal = retVal && actionEqual( lhsAction->MoveToFolder(), rhsAction->MoveToFolder() );
    retVal = retVal && actionEqual( lhsAction->NewItemAlert(), rhsAction->NewItemAlert() );
    retVal = retVal && actionEqual( lhsAction->PlaySound(), rhsAction->PlaySound() );
    retVal = retVal && actionEqual( lhsAction->ClearCategories(), rhsAction->ClearCategories() );
    retVal = retVal && actionEqual( lhsAction->Delete(), rhsAction->Delete() );
    retVal = retVal && actionEqual( lhsAction->DeletePermanently(), rhsAction->DeletePermanently() );
    retVal = retVal && actionEqual( lhsAction->DesktopAlert(), rhsAction->DesktopAlert() );
    retVal = retVal && actionEqual( lhsAction->NotifyDelivery(), rhsAction->NotifyDelivery() );
    retVal = retVal && actionEqual( lhsAction->NotifyRead(), rhsAction->NotifyRead() );
    retVal = retVal && actionEqual( lhsAction->Stop(), rhsAction->Stop() );
    retVal = retVal && actionEqual( lhsAction->CC(), rhsAction->CC() );
    retVal = retVal && actionEqual( lhsAction->Forward(), rhsAction->Forward() );
    retVal = retVal && actionEqual( lhsAction->ForwardAsAttachment(), rhsAction->ForwardAsAttachment() );
    retVal = retVal && actionEqual( lhsAction->Redirect(), rhsAction->Redirect() );
    return retVal;
}

bool conditionEqual( Outlook::AccountRuleCondition *lhsCondition, Outlook::AccountRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    if ( lhsCondition->ConditionType() != rhsCondition->ConditionType() )
        return false;

    return true;
}

bool conditionEqual( Outlook::RuleCondition *lhsCondition, Outlook::RuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    return true;
}

bool conditionEqual( Outlook::TextRuleCondition *lhsCondition, Outlook::TextRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    auto lhsText = toStringList( lhsCondition->Text() );
    auto rhsText = toStringList( rhsCondition->Text() );
    if ( !equal( lhsText, rhsText ) )
        return false;

    return true;
}

bool conditionEqual( Outlook::CategoryRuleCondition *lhsCondition, Outlook::CategoryRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;
    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    auto lhsText = toStringList( lhsCondition->Categories() );
    auto rhsText = toStringList( rhsCondition->Categories() );
    if ( !equal( lhsText, rhsText ) )
        return false;

    return true;
}

bool conditionEqual( Outlook::FormNameRuleCondition *lhsCondition, Outlook::FormNameRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;
    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    auto lhsText = toStringList( lhsCondition->FormName() );
    auto rhsText = toStringList( rhsCondition->FormName() );
    if ( !equal( lhsText, rhsText ) )
        return false;

    return true;
}

bool conditionEqual( Outlook::ToOrFromRuleCondition *lhsCondition, Outlook::ToOrFromRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    auto lhsAddresses = COutlookAPI::getEmailAddresses( lhsCondition->Recipients(), {}, false );
    auto rhsAddresses = COutlookAPI::getEmailAddresses( rhsCondition->Recipients(), {}, false );
    if ( !equal( lhsAddresses, rhsAddresses ) )
        return false;

    return true;
}

bool conditionEqual( Outlook::FromRssFeedRuleCondition *lhsCondition, Outlook::FromRssFeedRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    auto lhsText = toStringList( lhsCondition->FromRssFeed() );
    auto rhsText = toStringList( rhsCondition->FromRssFeed() );
    if ( !equal( lhsText, rhsText ) )
        return false;

    return true;
}

bool conditionEqual( Outlook::ImportanceRuleCondition *lhsCondition, Outlook::ImportanceRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    if ( lhsCondition->Importance() != rhsCondition->Importance() )
        return false;

    return true;
}

bool conditionEqual( Outlook::AddressRuleCondition *lhsCondition, Outlook::AddressRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    auto lhsText = toStringList( lhsCondition->Address() );
    auto rhsText = toStringList( rhsCondition->Address() );
    if ( !equal( lhsText, rhsText ) )
        return false;

    return true;
}

bool conditionEqual( Outlook::SenderInAddressListRuleCondition *lhsCondition, Outlook::SenderInAddressListRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    auto lhsAddresses = COutlookAPI::getEmailAddresses( lhsCondition->AddressList(), false );
    auto rhsAddresses = COutlookAPI::getEmailAddresses( rhsCondition->AddressList(), false );
    if ( !equal( lhsAddresses, rhsAddresses ) )
        return false;

    return true;
}

bool conditionEqual( Outlook::SensitivityRuleCondition *lhsCondition, Outlook::SensitivityRuleCondition *rhsCondition )
{
    if ( !lhsCondition || !rhsCondition )
        return false;

    if ( lhsCondition->Enabled() != rhsCondition->Enabled() )
        return false;

    if ( !lhsCondition->Enabled() || !rhsCondition->Enabled() )
        return true;

    if ( lhsCondition->Sensitivity() != rhsCondition->Sensitivity() )
        return false;

    return true;
}

std::optional< int > numConditionsDifferent( Outlook::RuleConditions *lhs, Outlook::RuleConditions *rhs )
{
    if ( !lhs || !rhs )
        return {};

    bool aOK = true;
    aOK = aOK && conditionEqual( lhs->Account(), rhs->Account() );
    aOK = aOK && conditionEqual( lhs->AnyCategory(), rhs->AnyCategory() );
    aOK = aOK && conditionEqual( lhs->CC(), rhs->CC() );
    aOK = aOK && conditionEqual( lhs->FromAnyRSSFeed(), rhs->FromAnyRSSFeed() );
    aOK = aOK && conditionEqual( lhs->Importance(), rhs->Importance() );
    aOK = aOK && conditionEqual( lhs->MeetingInviteOrUpdate(), rhs->MeetingInviteOrUpdate() );
    aOK = aOK && conditionEqual( lhs->NotTo(), rhs->NotTo() );
    aOK = aOK && conditionEqual( lhs->OnLocalMachine(), rhs->OnLocalMachine() );
    aOK = aOK && conditionEqual( lhs->OnOtherMachine(), rhs->OnOtherMachine() );
    aOK = aOK && conditionEqual( lhs->OnlyToMe(), rhs->OnlyToMe() );
    aOK = aOK && conditionEqual( lhs->Sensitivity(), rhs->Sensitivity() );
    aOK = aOK && conditionEqual( lhs->ToMe(), rhs->ToMe() );
    aOK = aOK && conditionEqual( lhs->ToOrCc(), rhs->ToOrCc() );
    if ( !aOK )
        return {};

    int retVal = 0;
    retVal += conditionEqual( lhs->Account(), rhs->Account() ) ? 0 : 1;
    retVal += conditionEqual( lhs->Body(), rhs->Body() ) ? 0 : 1;
    retVal += conditionEqual( lhs->BodyOrSubject(), rhs->BodyOrSubject() ) ? 0 : 1;
    retVal += conditionEqual( lhs->Category(), rhs->Category() ) ? 0 : 1;
    retVal += conditionEqual( lhs->FormName(), rhs->FormName() ) ? 0 : 1;
    retVal += conditionEqual( lhs->From(), rhs->From() ) ? 0 : 1;
    retVal += conditionEqual( lhs->FromRssFeed(), rhs->FromRssFeed() ) ? 0 : 1;
    retVal += conditionEqual( lhs->MessageHeader(), rhs->MessageHeader() ) ? 0 : 1;
    retVal += conditionEqual( lhs->RecipientAddress(), rhs->RecipientAddress() ) ? 0 : 1;
    retVal += conditionEqual( lhs->SenderAddress(), rhs->SenderAddress() ) ? 0 : 1;
    retVal += conditionEqual( lhs->SentTo(), rhs->SentTo() ) ? 0 : 1;
    retVal += conditionEqual( lhs->Subject(), rhs->Subject() ) ? 0 : 1;

    return retVal;
}
