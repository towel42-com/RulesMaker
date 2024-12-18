#include "OutlookAPI.h"

#include <QMessageBox>
#include "MSOUTL.h"
#include <tuple>

bool COutlookAPI::enableAllRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Enabling Rules:", numRules );

    std::list< Outlook::_Rule * > rules;
    int numChanged = 0;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;
        auto rule = fRules->Item( ii );
        emit sigIncStatusValue( "Enabling Rules:" );
        if ( !rule )
            continue;
        if ( rule->Enabled() )
            continue;
        emit sigStatusMessage( QString( "Enabling rule '%1'" ).arg( rule->Name() ) );
        rule->SetEnabled( true );
        numChanged++;
    }
    if ( canceled() )
        return false;

    if ( fParentWidget )
        QMessageBox::information( fParentWidget, R"(Enable All Rules)", QString( "%1 rules enabled" ).arg( numChanged ) );
    else
        emit sigStatusMessage( QString( "%1 rules enabled" ).arg( numChanged ) );

    if ( needsSaving )
        *needsSaving = numChanged != 0;

    if ( andSave && ( numChanged != 0 ) )
        saveRules();
    return true;
}

std::optional< QString > COutlookAPI::mergeKey( const std::shared_ptr< Outlook::Rule > &rule ) const
{
    auto from = rule->Conditions()->SenderAddress();
    if ( !from || !from->Enabled() )
        return {};

    auto moveAction = rule->Actions()->MoveToFolder();
    if ( !moveAction || !moveAction->Enabled() )
        return {};
    auto key = COutlookAPI::ruleNameForRule( rule );
    return key;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs )
{
    auto lhsRecipients = getRecipients( lhs, msgs );
    if ( !lhsRecipients.has_value() )
        return {};
    if ( lhsRecipients.value().empty() )
        return rhs;
    lhsRecipients.value() << rhs;
    lhsRecipients.value().removeDuplicates();
    return lhsRecipients;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs )
{
    auto tmpRhsRecipients = getRecipients( rhs, msgs );
    QStringList rhsRecipients;
    if ( tmpRhsRecipients.has_value() )
        rhsRecipients = tmpRhsRecipients.value();

    return mergeRecipients( lhs, rhsRecipients, msgs );
}

std::optional< QStringList > COutlookAPI::mergeRecipients( const std::list< Outlook::Rule * > &rules, QStringList *msgs )
{
    if ( rules.empty() )
        return {};
    auto pos = rules.begin();
    auto primaryRule = ( *pos );
    pos = std::next( pos );

    std::optional< QStringList > retVal;
    for ( ; pos != rules.end(); ++pos )
    {
        auto currRecipients = mergeRecipients( primaryRule, ( *pos ), msgs );
        if ( retVal.has_value() )
            retVal.value() << currRecipients.value();
        else
            retVal = currRecipients;
    }

    if ( retVal.has_value() )
    {
        retVal.value().removeAll( QString() );
        retVal.value().removeDuplicates();
    }

    return retVal;
}

bool COutlookAPI::mergeRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Analyzing Rules for Merging:", numRules );

    std::map< QString, std::list< std::pair< std::shared_ptr< Outlook::Rule >, int > > > rules;

    for ( int ii = 1; ii <= numRules; ++ii )
    {
        emit sigIncStatusValue( "Analyzing Rules for Merging:" );
        if ( canceled() )
            return false;

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule || !rule->Enabled() )
            continue;

        auto key = mergeKey( rule );
        if ( !key.has_value() )
            continue;

        rules[ key.value() ].emplace_back( rule, ii );
    }
    if ( canceled() )
        return false;

    bool forMsgBox = fParentWidget != nullptr;

    QString msg;
    for ( auto &&ii = rules.begin(); ii != rules.end(); )
    {
        if ( ( *ii ).second.size() < 2 )
        {
            ii = rules.erase( ii );
            continue;
        }
        auto &&curr = *ii;
        ++ii;

        if ( curr.second.size() < 2 )
            continue;

        auto pos = curr.second.begin();
        auto ruleName = ( *pos ).first->Name();
        if ( forMsgBox )
            ruleName = ruleName.toHtmlEscaped();

        QString currMsg = QString( "Primary Rule: %1 (%2)" ).arg( ruleName ).arg( ( *pos ).second );
        if ( forMsgBox )
            currMsg = "\t<li>" + currMsg + "\n\t\t<ul>";
        ++pos;
        for ( ; pos != curr.second.end(); ++pos )
        {
            ruleName = ( *pos ).first->Name();
            if ( forMsgBox )
                ruleName = ruleName.toHtmlEscaped();

            auto currRule = QString( "%1 (%2)" ).arg( ruleName ).arg( ( *pos ).second );
            if ( forMsgBox )
                currRule = "\n\t\t\t<li>" + currRule + "</li>";
            currRule += "\n";
            currMsg += currRule;
        }
        if ( forMsgBox )
            currMsg += "\t\t</ul>\n\t</li>";
        currMsg += "\n";
        msg += currMsg;
    }

    auto title = QString( "%1 merge(s) found%2" ).arg( rules.size() ).arg( forMsgBox ? "<br>\n" : "\n" );
    if ( !rules.empty() )
        title += QString( "Merging the following rules:\n%2" ).arg( forMsgBox ? "<ul>\n" : "" );

    msg = title + msg;
    if ( !forMsgBox )
        emit sigStatusMessage( msg );
    else
    {
        if ( !rules.empty() )
        {
            msg += "</ul>Do you wish to continue?";
            msg = msg.remove( "\n" );

            auto process = QMessageBox::information( fParentWidget, R"(Merge Rules by Target Folder)", msg, QMessageBox::Yes | QMessageBox::No );
            if ( process == QMessageBox::No )
                return false;
        }
        else
        {
            QMessageBox::information( fParentWidget, R"(Merge Rules by Target Folder)", msg );
            return false;
        }
    }

    if ( canceled() )
        return false;

    emit sigInitStatus( "Merging Rules:", static_cast< int >( rules.size() ) );

    std::list< int > toRemove;
    for ( auto &&ii : rules )
    {
        if ( canceled() )
            return false;

        std::list< Outlook::Rule * > currRules;
        for ( auto &&jj : ii.second )
            currRules.push_back( jj.first.get() );

        auto mergedRecipients = mergeRecipients( currRules, nullptr );
        if ( !mergedRecipients.has_value() )
            continue;

        auto pos = std::next( ii.second.begin() );

        for ( ; pos != ii.second.end(); ++pos )
        {
            if ( disableRatherThanDeleteRules() )
                ( *pos ).first->SetEnabled( false );
            else
                toRemove.push_back( ( *pos ).second );
        }
        ii.second.front().first->Conditions()->SenderAddress()->SetAddress( mergedRecipients.value() );

        emit sigIncStatusValue( "Merging Rules:" );
    }

    toRemove.sort();
    if ( !toRemove.empty() )
        emit sigInitStatus( "Deleting old Rules:", static_cast< int >( toRemove.size() ) );

    for ( auto &&ii = toRemove.rbegin(); ii != toRemove.rend(); ++ii )
    {
        if ( canceled() )
            return false;
        fRules->Remove( *ii );
        emit sigIncStatusValue( "Deleting old Rules:" );
    }

    if ( needsSaving )
        *needsSaving = !rules.empty();

    if ( andSave && !rules.empty() )
        saveRules();

    return true;
}

bool COutlookAPI::moveFromToAddress( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Fixing Rules:", numRules );
    int numChanged = 0;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        emit sigIncStatusValue( "Fixing Rules:" );

        auto conditions = rule->Conditions();
        if ( !conditions )
            continue;

        auto from = conditions->From();
        if ( !from->Enabled() )
            continue;

        emit sigStatusMessage( QString( "Checking from email addresses on rule '%1'" ).arg( rule->Name() ) );
        auto fromEmails = getEmailAddresses( from->Recipients(), {}, true );
        if ( fromEmails.isEmpty() )
            continue;

        QStringList msgs;
        if ( !addRecipientsToRule( rule.get(), fromEmails, msgs ) )
            return false;

        from->SetEnabled( false );
        numChanged++;
    }
    if ( numChanged )
        saveRules();

    if ( fParentWidget )
        QMessageBox::information( fParentWidget, R"(Move "From" to "Address")", QString( "%1 rules modified" ).arg( numChanged ) );
    else
        emit sigStatusMessage( QString( "%1 rules modified" ).arg( numChanged ) );

    if ( needsSaving )
        *needsSaving = numChanged != 0;

    if ( andSave && ( numChanged != 0 ) )
        saveRules();

    return true;
}

bool COutlookAPI::renameRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();
    auto numRules = fRules->Count();

    emit sigInitStatus( "Analyzing Rule Names:", numRules );

    std::list< std::pair< std::shared_ptr< Outlook::Rule >, QString > > changes;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;

        emit sigIncStatusValue( "Analyzing Rule Names:" );
        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        auto ruleName = ruleNameForRule( rule );
        auto currName = rule->Name();
        if ( ruleName != currName )
        {
            changes.emplace_back( rule, ruleName );
        }
    }
    if ( canceled() )
        return false;

    if ( changes.empty() )
    {
        if ( fParentWidget )
            QMessageBox::information( fParentWidget, "Renamed Rules", QString( "No rules needed renaming" ) );
        else
            emit sigStatusMessage( QString( "No rules needed renaming" ) );
        return true;
    }
    if ( fParentWidget )
    {
        QStringList tmp;
        for ( auto &&ii : changes )
        {
            tmp << "<li>" + ii.first->Name().toHtmlEscaped() + " => " + ii.second.toHtmlEscaped() + "</li>";
        }
        auto msg = QString( "Rules to be changed:<ul>%1</ul>Do you wish to continue?" ).arg( tmp.join( "\n" ) );
        auto process = QMessageBox::information( fParentWidget, "Renamed Rules", msg, QMessageBox::Yes | QMessageBox::No );
        if ( process == QMessageBox::No )
            return false;
    }
    else
    {
        for ( auto &&ii : changes )
        {
            emit sigStatusMessage( QString( "Rule '%1' will be renamed to '%2'" ).arg( ii.first->Name(), ii.second ) );
        }
    }

    emit sigInitStatus( "Renaming Rules:", static_cast< int >( changes.size() ) );
    for ( auto &&ii : changes )
    {
        ii.first->SetName( ii.second );
        emit sigIncStatusValue( "Renaming Rules:" );
    }

    if ( needsSaving )
        *needsSaving = !changes.empty();

    if ( andSave && !changes.empty() )
        saveRules();
    saveRules();

    return true;
}

bool COutlookAPI::sortRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Sorting Rules:", numRules );

    std::list< Outlook::_Rule * > rules;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;
        auto rule = fRules->Item( ii );
        emit sigIncStatusValue( "Sorting Rules:" );
        if ( !rule )
            continue;
        rules.push_back( rule );
    }
    if ( canceled() )
        return false;

    rules.sort(
        []( Outlook::_Rule *lhs, Outlook::_Rule *rhs )
        {
            if ( !lhs )
                return false;
            if ( !rhs )
                return true;
            auto lhsName = lhs->Name();
            auto rhsName = rhs->Name();
            if ( lhsName.startsWith( rhsName ) && ( lhsName != rhsName ) )
                return true;
            else if ( rhsName.startsWith( lhsName ) && ( lhsName != rhsName ) )
                return false;
            else
                return lhsName < rhsName;
        } );

    if ( canceled() )
        return false;
    auto pos = 1;
    emit sigInitStatus( "Recomputing Execution Order:", numRules );

    std::list< std::tuple< QString, Outlook::_Rule *, int > > rulesChanged;
    for ( auto &&ii : rules )
    {
        if ( canceled() )
            return false;

        if ( ii->ExecutionOrder() != pos )
        {
            auto msg = QString( "%1 - %2 -> %3" ).arg( ii->Name() ).arg( ii->ExecutionOrder() ).arg( pos );
            if ( !fParentWidget )
                emit sigStatusMessage( msg );
            rulesChanged.emplace_back( msg.toHtmlEscaped(), ii, pos );
        }
        pos++;
        emit sigIncStatusValue( "Recomputing Execution Order:" );
    }

    auto msg = QString( "%1 rules needed re-ordering" ).arg( rulesChanged.size() );
    if ( fParentWidget )
    {
        if ( rulesChanged.empty() )
            QMessageBox::information( fParentWidget, R"(Sorting Rules by Name)", msg );
        else
        {
            msg += "<ul>";
            int cnt = 0;
            for ( auto &&ii : rulesChanged )
            {
                if ( cnt >= 5 )
                {
                    msg += "And More...";
                    break;
                }
                msg += "<li>" + std::get< 0 >( ii ) + "</li>";
                cnt++;
            }
            msg += "</ul>";
            msg += "Do you wish to Continue?";
            auto process = QMessageBox::information( fParentWidget, R"(Sorting Rules by Name)", msg, QMessageBox::Yes | QMessageBox::No );
            if ( process == QMessageBox::No )
                return false;
        }
    }
    else
        emit sigStatusMessage( msg );

    for ( auto &&ii : rulesChanged )
    {
        auto rule = std::get< 1 >( ii );
        rule->SetExecutionOrder( std::get< 2 >( ii ) );
    }

    if ( needsSaving )
        *needsSaving = rulesChanged.size() != 0;

    if ( andSave && ( rulesChanged.size() != 0 ) )
        saveRules();

    return true;
}

void COutlookAPI::slotHandleRulesSaveException( int, const QString &, const QString &, const QString & )
{
    fSaveRulesSuccess = false;
}

bool COutlookAPI::saveRules()
{
    emit sigStatusMessage( QString( "Saving Rules" ) );
    fSaveRulesSuccess = true;
    connect( fRules.get(), SIGNAL( exception( int, QString, QString, QString ) ), this, SLOT( slotHandleRulesSaveException( int, const QString &, const QString &, const QString & ) ) );
    fRules->Save( fParentWidget ? true : false );
    disconnect( fRules.get(), SIGNAL( exception( int, QString, QString, QString ) ), this, SLOT( slotHandleRulesSaveException( int, const QString &, const QString &, const QString & ) ) );
    return fSaveRulesSuccess;
}
