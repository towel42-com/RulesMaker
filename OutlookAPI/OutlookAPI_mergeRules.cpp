#include "OutlookAPI.h"
#include "OutlookAPI_pri.h"

#include "MSOUTL.h"
#include <map>
#include <QDebug>

std::optional< QString > COutlookAPI::mergeKey( const std::shared_ptr< Outlook::Rule > &rule ) const
{
    return getDestFolderNameForRule( rule, true );
}

bool COutlookAPI::canMergeRules( std::shared_ptr< Outlook::Rule > lhs, std::shared_ptr< Outlook::Rule > rhs )
{
    if ( !lhs || !rhs )
        return false;

    qDebug() << "Comparing: " << getDebugName( lhs ) << " to " << getDebugName( rhs );

    if ( !actionsEqual( lhs->Actions(), rhs->Actions() ) )
        return false;

    auto numConditonsDiff = numConditionsDifferent( lhs->Conditions(), rhs->Conditions() );
    if ( !numConditonsDiff.has_value() || ( numConditonsDiff.value() > 1 ) )
        return false;

    auto numExceptionsDiff = numConditionsDifferent( lhs->Exceptions(), rhs->Exceptions() );
    if ( !numExceptionsDiff.has_value() || ( numExceptionsDiff.value() > 1 ) )
        return false;

    if ( ( numConditonsDiff.value() + numExceptionsDiff.value() ) > 1 )
        return false;

    return true;
}

void mergeCondition( Outlook::TextRuleCondition *lhsCondition, Outlook::TextRuleCondition *rhsCondition )
{
    if ( conditionEqual( lhsCondition, rhsCondition ) )
        return;

    auto lhsText = lhsCondition->Enabled() ? toStringList( lhsCondition->Text() ) : QStringList();
    auto rhsText = rhsCondition->Enabled() ? toStringList( rhsCondition->Text() ) : QStringList();

    auto newText = mergeStringLists( lhsText, rhsText, true );
    lhsCondition->SetText( newText );
    lhsCondition->SetEnabled( rhsCondition->Enabled() );
}

void mergeCondition( Outlook::CategoryRuleCondition *lhsCondition, Outlook::CategoryRuleCondition *rhsCondition )
{
    if ( conditionEqual( lhsCondition, rhsCondition ) )
        return;

    auto lhsText = lhsCondition->Enabled() ? toStringList( lhsCondition->Categories() ) : QStringList();
    auto rhsText = rhsCondition->Enabled() ? toStringList( rhsCondition->Categories() ) : QStringList();

    auto newText = mergeStringLists( lhsText, rhsText, true );
    lhsCondition->SetCategories( newText );
    lhsCondition->SetEnabled( rhsCondition->Enabled() );
}

void mergeCondition( Outlook::FormNameRuleCondition *lhsCondition, Outlook::FormNameRuleCondition *rhsCondition )
{
    if ( conditionEqual( lhsCondition, rhsCondition ) )
        return;

    auto lhsText = lhsCondition->Enabled() ? toStringList( lhsCondition->FormName() ) : QStringList();
    auto rhsText = rhsCondition->Enabled() ? toStringList( rhsCondition->FormName() ) : QStringList();

    auto newText = mergeStringLists( lhsText, rhsText, true );
    lhsCondition->SetFormName( newText );
    lhsCondition->SetEnabled( rhsCondition->Enabled() );
}

void mergeCondition( Outlook::ToOrFromRuleCondition *lhsCondition, Outlook::ToOrFromRuleCondition *rhsCondition )
{
    if ( conditionEqual( lhsCondition, rhsCondition ) )
        return;

    auto lhsAddresses = lhsCondition->Enabled() ? COutlookAPI::getEmailAddresses( lhsCondition->Recipients(), {}, false ) : QStringList();
    auto rhsAddresses = rhsCondition->Enabled() ? COutlookAPI::getEmailAddresses( rhsCondition->Recipients(), {}, false ) : QStringList();
    auto newText = mergeStringLists( lhsAddresses, rhsAddresses, false );

    for ( auto &&ii : newText )
    {
        lhsCondition->Recipients()->Add( ii );
    }
    lhsCondition->SetEnabled( rhsCondition->Enabled() );
}

void mergeCondition( Outlook::FromRssFeedRuleCondition *lhsCondition, Outlook::FromRssFeedRuleCondition *rhsCondition )
{
    if ( conditionEqual( lhsCondition, rhsCondition ) )
        return;

    auto lhsText = lhsCondition->Enabled() ? toStringList( lhsCondition->FromRssFeed() ) : QStringList();
    auto rhsText = rhsCondition->Enabled() ? toStringList( rhsCondition->FromRssFeed() ) : QStringList();

    auto newText = mergeStringLists( lhsText, rhsText, true );
    lhsCondition->SetFromRssFeed( newText );
    lhsCondition->SetEnabled( rhsCondition->Enabled() );
}

void mergeCondition( Outlook::AddressRuleCondition *lhsCondition, Outlook::AddressRuleCondition *rhsCondition )
{
    if ( conditionEqual( lhsCondition, rhsCondition ) )
        return;

    auto lhsText = lhsCondition->Enabled() ? toStringList( lhsCondition->Address() ) : QStringList();
    auto rhsText = rhsCondition->Enabled() ? toStringList( rhsCondition->Address() ) : QStringList();

    auto newText = mergeStringLists( lhsText, rhsText, true );
    lhsCondition->SetAddress( newText );
    lhsCondition->SetEnabled( rhsCondition->Enabled() );
}

void mergeConditions( Outlook::RuleConditions *lhs, Outlook::RuleConditions *rhs )
{
    if ( !lhs || !rhs )
        return;

    mergeCondition( lhs->Body(), rhs->Body() );
    mergeCondition( lhs->BodyOrSubject(), rhs->BodyOrSubject() );
    mergeCondition( lhs->Category(), rhs->Category() );
    mergeCondition( lhs->FormName(), rhs->FormName() );
    mergeCondition( lhs->From(), rhs->From() );
    mergeCondition( lhs->FromRssFeed(), rhs->FromRssFeed() );
    mergeCondition( lhs->MessageHeader(), rhs->MessageHeader() );
    mergeCondition( lhs->RecipientAddress(), rhs->RecipientAddress() );
    mergeCondition( lhs->SenderAddress(), rhs->SenderAddress() );
    mergeCondition( lhs->SentTo(), rhs->SentTo() );
    mergeCondition( lhs->Subject(), rhs->Subject() );
}

std::shared_ptr< Outlook::Rule > COutlookAPI::mergeRule( std::shared_ptr< Outlook::Rule > &lhs, std::shared_ptr< Outlook::Rule > &rhs )
{
    if ( !canMergeRules( lhs, rhs ) )
        return {};

    if ( !lhs || !rhs )
        return {};

    qDebug() << "Merging: " << getDebugName( rhs ) << " to " << getDebugName( lhs );

    mergeConditions( lhs->Conditions(), rhs->Conditions() );
    mergeConditions( lhs->Exceptions(), rhs->Exceptions() );

    return lhs;
}

void COutlookAPI::mergeRules( COutlookAPI::TRulePair &rules )
{
    if ( !rules.first || rules.second.empty() )
        return;

    auto && primaryRule = rules.first;
    qDebug() << "Primary Rule: " << getDebugName( primaryRule );

    auto && toRemove = rules.second;
    for ( auto &&ii : toRemove )
    {
        qDebug() << "    To be Merged: " << getDebugName( ii );
    }

    for ( auto &&ii = toRemove.begin(); ii != toRemove.end(); )
    {
        if ( !mergeRule( primaryRule, *ii ) )
        {
            ii = toRemove.erase( ii );
            return;
        }
        deleteRule( *ii, true, false );
        ++ii;
    }
    auto ruleName = ruleNameForRule( primaryRule );
    primaryRule->SetName( ruleName );
}

COutlookAPI::TMergeRuleMap COutlookAPI::findMergableRules()
{
    TMergeRuleMap retVal;
    auto numRules = fRules->Count();
    emit sigInitStatus( "Analyzing Rules for Merging:", numRules );
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        emit sigIncStatusValue( "Analyzing Rules for Merging:" );
        if ( canceled() )
            return {};

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule || !rule->Enabled() )
            continue;

        auto key = mergeKey( rule );
        if ( !key.has_value() )
            continue;

        auto range = retVal.equal_range( key.value() );
        bool matched = false;
        for ( auto &&pos = range.first; pos != range.second; ++pos )
        {
            auto primary = ( *pos ).second.first;
            if ( !canMergeRules( primary, rule ) )
                continue;
            ( *pos ).second.second.push_back( rule );
            matched = true;
            break;
        }
        if ( !matched )
        {
            retVal.emplace( key.value(), std::make_pair( rule, TRuleList() ) );
        }
    }
    if ( canceled() )
        return {};

    for ( auto &&ii = retVal.begin(); ii != retVal.end(); )
    {
        if ( ( *ii ).second.first )
        {
            for ( auto &&jj = ( *ii ).second.second.begin(); jj != ( *ii ).second.second.end(); )
            {
                if ( !( *jj ) )
                {
                    jj = ( *ii ).second.second.erase( jj );
                    continue;
                }
                ++jj;
            }
        }

        if ( !( *ii ).second.first || ( *ii ).second.second.empty() )
        {
            ii = retVal.erase( ii );
            continue;
        }
        ++ii;
    }

    for ( auto &&ii : retVal )
    {
        ii.second.second.sort( []( const std::shared_ptr< Outlook::Rule > &lhs, const std::shared_ptr< Outlook::Rule > &rhs ) { return lhs->ExecutionOrder() > rhs->ExecutionOrder(); } );
    }

    return retVal;
}