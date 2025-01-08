#include "OutlookAPI.h"
#include "MSOUTL.h"

void copyAction( Outlook::AssignToCategoryRuleAction *lhsAction, Outlook::AssignToCategoryRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return;

    lhsAction->SetEnabled( rhsAction->Enabled() );
    if ( !rhsAction->Enabled() )
        return;

    lhsAction->SetCategories( rhsAction->Categories() );
}

void copyAction( Outlook::MarkAsTaskRuleAction *lhsAction, Outlook::MarkAsTaskRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return;

    lhsAction->SetEnabled( rhsAction->Enabled() );
    if ( !rhsAction->Enabled() )
        return;

    lhsAction->SetFlagTo( rhsAction->FlagTo() );
    lhsAction->SetMarkInterval( rhsAction->MarkInterval() );
}

void copyAction( Outlook::MoveOrCopyRuleAction *lhsAction, Outlook::MoveOrCopyRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return;
    lhsAction->SetEnabled( rhsAction->Enabled() );
    if ( !rhsAction->Enabled() )
        return;

    lhsAction->SetFolder( rhsAction->Folder() );
}

void copyAction( Outlook::NewItemAlertRuleAction *lhsAction, Outlook::NewItemAlertRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return;

    lhsAction->SetEnabled( rhsAction->Enabled() );
    if ( !rhsAction->Enabled() )
        return;

    lhsAction->SetText( rhsAction->Text() );
}

void copyAction( Outlook::PlaySoundRuleAction *lhsAction, Outlook::PlaySoundRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return;

    lhsAction->SetEnabled( rhsAction->Enabled() );
    if ( !rhsAction->Enabled() )
        return;

    lhsAction->SetFilePath( rhsAction->FilePath() );
}

void copyAction( Outlook::RuleAction *lhsAction, Outlook::RuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return;

    lhsAction->SetEnabled( rhsAction->Enabled() );
    if ( !rhsAction->Enabled() )
        return;
}

void copyAction( Outlook::SendRuleAction *lhsAction, Outlook::SendRuleAction *rhsAction )
{
    if ( !lhsAction || !rhsAction )
        return;

    lhsAction->SetEnabled( rhsAction->Enabled() );
    if ( !rhsAction->Enabled() )
        return;

    auto rhsAddresses = rhsAction->Enabled() ? COutlookAPI::getEmailAddresses( rhsAction->Recipients(), {}, false ) : QStringList();
    for ( auto &&ii : rhsAddresses )
    {
        lhsAction->Recipients()->Add( ii );
    }
}

void copyActions( std::shared_ptr< Outlook::Rule > retValRule, std::shared_ptr< Outlook::Rule > source )
{
    if ( !retValRule || !source )
        return;

    auto retVal = retValRule->Actions();
    auto sourceActions = source->Actions();

    copyAction( retVal->AssignToCategory(), sourceActions->AssignToCategory() );
    copyAction( retVal->MarkAsTask(), sourceActions->MarkAsTask() );
    copyAction( retVal->CopyToFolder(), sourceActions->CopyToFolder() );
    copyAction( retVal->MoveToFolder(), sourceActions->MoveToFolder() );
    copyAction( retVal->NewItemAlert(), sourceActions->NewItemAlert() );
    copyAction( retVal->PlaySound(), sourceActions->PlaySound() );
    copyAction( retVal->ClearCategories(), sourceActions->ClearCategories() );
    copyAction( retVal->Delete(), sourceActions->Delete() );
    copyAction( retVal->DeletePermanently(), sourceActions->DeletePermanently() );
    copyAction( retVal->DesktopAlert(), sourceActions->DesktopAlert() );
    copyAction( retVal->NotifyDelivery(), sourceActions->NotifyDelivery() );
    copyAction( retVal->NotifyRead(), sourceActions->NotifyRead() );
    copyAction( retVal->Stop(), sourceActions->Stop() );
    copyAction( retVal->CC(), sourceActions->CC() );
    copyAction( retVal->Forward(), sourceActions->Forward() );
    copyAction( retVal->ForwardAsAttachment(), sourceActions->ForwardAsAttachment() );
    copyAction( retVal->Redirect(), sourceActions->Redirect() );
}

void copyCondition( Outlook::AccountRuleCondition *retVal, Outlook::AccountRuleCondition *sourceCondition );
void copyCondition( Outlook::RuleCondition *retVal, Outlook::RuleCondition *sourceCondition );
void copyCondition( Outlook::TextRuleCondition *retVal, Outlook::TextRuleCondition *sourceCondition );
void copyCondition( Outlook::CategoryRuleCondition *retVal, Outlook::CategoryRuleCondition *sourceCondition );
void copyCondition( Outlook::ToOrFromRuleCondition *retVal, Outlook::ToOrFromRuleCondition *sourceCondition );
void copyCondition( Outlook::FormNameRuleCondition *retVal, Outlook::FormNameRuleCondition *sourceCondition );
void copyCondition( Outlook::FromRssFeedRuleCondition *retVal, Outlook::FromRssFeedRuleCondition *sourceCondition );
void copyCondition( Outlook::ImportanceRuleCondition *retVal, Outlook::ImportanceRuleCondition *sourceCondition );
void copyCondition( Outlook::AddressRuleCondition *retVal, Outlook::AddressRuleCondition *sourceCondition );
void copyCondition( Outlook::SenderInAddressListRuleCondition *retVal, Outlook::SenderInAddressListRuleCondition *sourceCondition );
void copyCondition( Outlook::SensitivityRuleCondition *retVal, Outlook::SensitivityRuleCondition *sourceCondition );

void copyCondition( Outlook::AccountRuleCondition *retVal, Outlook::AccountRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;
}

void copyCondition( Outlook::RuleCondition *retVal, Outlook::RuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;
}

void copyCondition( Outlook::TextRuleCondition *retVal, Outlook::TextRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    retVal->SetText( sourceCondition->Text() );
}

void copyCondition( Outlook::CategoryRuleCondition *retVal, Outlook::CategoryRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    retVal->SetCategories( sourceCondition->Categories() );
}

void copyCondition( Outlook::FormNameRuleCondition *retVal, Outlook::FormNameRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    retVal->SetFormName( sourceCondition->FormName() );
}

void copyCondition( Outlook::ToOrFromRuleCondition *retVal, Outlook::ToOrFromRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    auto rhsAddresses = sourceCondition->Enabled() ? COutlookAPI::getEmailAddresses( sourceCondition->Recipients(), {}, false ) : QStringList();
    for ( auto &&ii : rhsAddresses )
    {
        retVal->Recipients()->Add( ii );
    }
}

void copyCondition( Outlook::FromRssFeedRuleCondition *retVal, Outlook::FromRssFeedRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    retVal->SetFromRssFeed( sourceCondition->FromRssFeed() );
}

void copyCondition( Outlook::ImportanceRuleCondition *retVal, Outlook::ImportanceRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;
    retVal->SetImportance( sourceCondition->Importance() );
}

void copyCondition( Outlook::AddressRuleCondition *retVal, Outlook::AddressRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    retVal->SetAddress( sourceCondition->Address() );
}

void copyCondition( Outlook::SenderInAddressListRuleCondition *retVal, Outlook::SenderInAddressListRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    retVal->SetAddressList( sourceCondition->AddressList() );
}

void copyCondition( Outlook::SensitivityRuleCondition *retVal, Outlook::SensitivityRuleCondition *sourceCondition )
{
    if ( !retVal || !sourceCondition )
        return;

    retVal->SetEnabled( sourceCondition->Enabled() );
    if ( !sourceCondition->Enabled() )
        return;

    retVal->SetSensitivity( sourceCondition->Sensitivity() );
}

void copyConditions( std::shared_ptr< Outlook::Rule > retValRule, std::shared_ptr< Outlook::Rule > source, bool exceptions )
{
    if ( !retValRule || !source )
        return;

    auto retVal = exceptions ? retValRule->Exceptions() : retValRule->Conditions();
    auto sourceConditions = exceptions ? source->Exceptions() : source->Conditions();

    copyCondition( retVal->Account(), sourceConditions->Account() );
    copyCondition( retVal->AnyCategory(), sourceConditions->AnyCategory() );
    copyCondition( retVal->Body(), sourceConditions->Body() );
    copyCondition( retVal->BodyOrSubject(), sourceConditions->BodyOrSubject() );
    copyCondition( retVal->CC(), sourceConditions->CC() );
    copyCondition( retVal->Category(), sourceConditions->Category() );
    copyCondition( retVal->FormName(), sourceConditions->FormName() );
    copyCondition( retVal->From(), sourceConditions->From() );
    copyCondition( retVal->FromAnyRSSFeed(), sourceConditions->FromAnyRSSFeed() );
    copyCondition( retVal->FromRssFeed(), sourceConditions->FromRssFeed() );
    copyCondition( retVal->HasAttachment(), sourceConditions->HasAttachment() );
    copyCondition( retVal->Importance(), sourceConditions->Importance() );
    copyCondition( retVal->MeetingInviteOrUpdate(), sourceConditions->MeetingInviteOrUpdate() );
    copyCondition( retVal->MessageHeader(), sourceConditions->MessageHeader() );
    copyCondition( retVal->NotTo(), sourceConditions->NotTo() );
    copyCondition( retVal->OnLocalMachine(), sourceConditions->OnLocalMachine() );
    copyCondition( retVal->OnOtherMachine(), sourceConditions->OnOtherMachine() );
    copyCondition( retVal->OnlyToMe(), sourceConditions->OnlyToMe() );
    copyCondition( retVal->RecipientAddress(), sourceConditions->RecipientAddress() );
    copyCondition( retVal->SenderAddress(), sourceConditions->SenderAddress() );
    copyCondition( retVal->SenderInAddressList(), sourceConditions->SenderInAddressList() );
    copyCondition( retVal->Sensitivity(), sourceConditions->Sensitivity() );
    copyCondition( retVal->SentTo(), sourceConditions->SentTo() );
    copyCondition( retVal->Subject(), sourceConditions->Subject() );
    copyCondition( retVal->ToMe(), sourceConditions->ToMe() );
    copyCondition( retVal->ToOrCc(), sourceConditions->ToOrCc() );
}

std::shared_ptr< Outlook::Rule > COutlookAPI::copyRule( std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return {};

    auto retVal = std::shared_ptr< Outlook::Rule >( fRules->Create( rule->Name(), rule->RuleType() ) );
    if ( !retVal )
        return {};

    retVal->SetEnabled( rule->Enabled() );

    copyActions( retVal, rule );
    copyConditions( retVal, rule, false );
    copyConditions( retVal, rule, true );
    return retVal;
}
