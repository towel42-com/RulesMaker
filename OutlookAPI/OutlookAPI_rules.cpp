#include "OutlookAPI.h"

#include <QMessageBox>
#include <QRegularExpression>

#include "MSOUTL.h"

std::pair< std::shared_ptr< Outlook::Rules >, int > COutlookAPI::getRules()
{
    if ( !fRules )
        fRules = selectRules();
    return { fRules, fRules->Count() };
}

std::shared_ptr< Outlook::Rule > COutlookAPI::getRule( const std::shared_ptr< Outlook::Rules > &rules, int num )
{
    if ( !rules || !num || ( num > rules->Count() ) )
        return {};
    auto rule = rules->Item( num );
    if ( !rule )
        return {};
    return getRule( rule );
}

std::optional< bool > COutlookAPI::addRule( const std::shared_ptr< Outlook::Folder > &folder, const std::list< std::pair< QStringList, EFilterType > > &patterns, QStringList &msgs )
{
    if ( !folder )
    {
        msgs.push_back( "Parameters not set" );
        return false;
    }

    auto ruleName = ruleNameForFolder( folder );

    auto rule = std::shared_ptr< Outlook::Rule >( fRules->Create( ruleName, Outlook::OlRuleType::olRuleReceive ) );
    if ( !rule )
    {
        msgs.push_back( QString( "Could not create rule '%1'" ).arg( ruleName ) );
        return false;
    }

    auto moveAction = rule->Actions()->MoveToFolder();
    if ( !moveAction )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }
    moveAction->SetEnabled( true );
    moveAction->SetFolder( reinterpret_cast< Outlook::MAPIFolder * >( folder.get() ) );

    rule->Actions()->Stop()->SetEnabled( true );

    return addToRule( rule, patterns, msgs, false );
}

std::optional< bool > COutlookAPI::addToRule( std::shared_ptr< Outlook::Rule > rule, const std::list< std::pair< QStringList, EFilterType > > &patterns, QStringList &msgs, bool copyFirst )
{
    bool patternsEmpty = patterns.empty();
    for ( auto &&ii = patterns.begin(); !patternsEmpty && ( ii != patterns.end() ); ++ii )
    {
        patternsEmpty = patternsEmpty || ii->first.empty();
    }

    if ( !rule || patternsEmpty || !fRules )
    {
        msgs.push_back( "Parameters not set" );
        return false;
    }
    auto newRule = copyFirst ? copyRule( rule ) : rule;
    if ( !newRule )
    {
        msgs.push_back( "Could not backup rule" );
        return false;
    }

    for ( auto &&[ rules, patternType ] : patterns )
    {
        switch ( patternType )
        {
            case EFilterType::eByEmailAddress:
                {
                    if ( !addRecipientsToRule( newRule.get(), rules, msgs ) )
                        return false;
                }
                break;
            case EFilterType::eByDisplayName:
                {
                    if ( !addDisplayNamesToRule( newRule.get(), rules, msgs ) )
                        return false;
                }
                break;
            case EFilterType::eBySubject:
                {
                    if ( !addSubjectsToRule( newRule.get(), rules, msgs ) )
                        return false;
                }
                break;
            case EFilterType::eByOutlookContact:
                {
                    if ( !addOutlookContactsToRule( newRule.get(), rules, msgs ) )
                        return false;
                }
                break;
            default:
                return false;
        }
    }
    auto name = ruleNameForRule( newRule, false );
    if ( newRule->Name() != name )
        newRule->SetName( name );

    if ( !showRule( newRule ) )
        return {};

    if ( copyFirst )
    {
        auto executionOrder = rule->ExecutionOrder();
        if ( !deleteRule( rule, false, false ) )
        {
            msgs.push_back( "Could not delete original rule" );
            return false;
        }
        newRule->SetExecutionOrder( executionOrder );
    }
    saveRules();

    bool retVal = runRule( newRule );
    if ( !retVal )
    {
        msgs.push_back( "Could not run rule, but it was created" );
    }
    emit sigRuleChanged( newRule );
    return retVal;
}

bool COutlookAPI::ruleEnabled( const std::shared_ptr< Outlook::Rule > &rule )
{
    if ( !rule )
        return false;
    return rule->Enabled();
}

bool COutlookAPI::deleteRule( std::shared_ptr< Outlook::Rule > rule, bool forceDisable, bool andSave )
{
    if ( !rule || !fRules )
        return false;

    auto ruleName = getDisplayName( rule );

    auto disable = forceDisable || disableRatherThanDeleteRules();
    emit sigStatusMessage( QString( "%1 Rule: %2" ).arg( ( disable ? "Disabling" : "Deleting" ), ruleName ) );
    bool aOK = false;
    if ( disable )
        aOK = disableRule( rule, andSave );
    else
    {
        fRules->Remove( rule->ExecutionOrder() );
        aOK = true;
    }
    if ( !aOK )
        return false;

    auto pos = fRuleBeenLoaded.find( rule );
    if ( pos != fRuleBeenLoaded.end() )
    {
        fRuleBeenLoaded.erase( pos );
    }

    if ( andSave )
    {
        saveRules();
        QMessageBox::information( fParentWidget, "Deleted Rule", QString( "Deleted Rule: %1" ).arg( ruleName ) );
    }

    if ( !disable )
        emit sigRuleDeleted( rule );
    return true;
}

bool COutlookAPI::disableRule( const std::shared_ptr< Outlook::Rule > &rule, bool andSave )
{
    auto ruleName = getDisplayName( rule );

    emit sigStatusMessage( QString( "Disabling Rule: %1" ).arg( ruleName ) );
    rule->SetEnabled( false );

    if ( andSave )
    {
        saveRules();
        QMessageBox::information( fParentWidget, "Disabled Rule", QString( "Disabled Rule: %1" ).arg( ruleName ) );
    }
    emit sigRuleChanged( rule );
    return true;
}

bool COutlookAPI::enableRule( const std::shared_ptr< Outlook::Rule > &rule, bool andSave )
{
    auto ruleName = getDisplayName( rule );

    emit sigStatusMessage( QString( "Disabling Rule: %1" ).arg( ruleName ) );
    rule->SetEnabled( true );

    if ( andSave )
    {
        saveRules();
        QMessageBox::information( fParentWidget, "Enabled Rule", QString( "Enabled Rule: %1" ).arg( ruleName ) );
    }

    emit sigRuleChanged( rule );
    return true;
}

QString COutlookAPI::moveTargetFolderForRule( const std::shared_ptr< Outlook::Rule > &rule ) const
{
    if ( !rule )
        return {};
    auto moveAction = rule->Actions()->MoveToFolder();
    if ( !moveAction || !moveAction->Enabled() || !moveAction->Folder() )
        return {};

    auto folderName = moveAction->Folder()->FolderPath();
    return folderName;
}

std::list< EFilterType > COutlookAPI::filterTypesForRule( const std::shared_ptr< Outlook::Rule > &rule ) const
{
    if ( !rule )
        return {};

    std::list< EFilterType > retVal;
    auto conditions = rule->Conditions();
    if ( !conditions )
        return retVal;
    auto senderAddress = conditions->SenderAddress();
    if ( senderAddress && senderAddress->Enabled() )
        retVal.push_back( EFilterType::eByEmailAddress );

    auto header = rule->Conditions()->MessageHeader();
    if ( header && header->Enabled() )
        retVal.push_back( EFilterType::eByDisplayName );

    auto subject = rule->Conditions()->Subject();
    if ( subject && subject->Enabled() )
        retVal.push_back( EFilterType::eBySubject );

    //auto outlookContact = rule->Conditions()->Subject();
    //if ( outlookContact && outlookContact->Enabled() )
    //    retVal.push_back( EFilterType::eByOutlookContact );
    return retVal;
}

bool COutlookAPI::isEnabled( const std::shared_ptr< Outlook::Rule > &rule )
{
    if ( !rule )
        return false;
    return rule->Enabled();
}

bool COutlookAPI::ruleBeenLoaded( std::shared_ptr< Outlook::Rule > &rule ) const
{
    auto pos = fRuleBeenLoaded.find( rule );
    return pos != fRuleBeenLoaded.end();
}

bool COutlookAPI::ruleLessThan( const std::shared_ptr< Outlook::Rule > &lhsRule, const std::shared_ptr< Outlook::Rule > &rhsRule ) const
{
    if ( !lhsRule )
        return false;
    if ( !rhsRule )
        return true;
    return lhsRule->ExecutionOrder() < rhsRule->ExecutionOrder();
}

bool COutlookAPI::runAllRules( std::shared_ptr< Outlook::Folder > folder, bool allFolders, bool junk )
{
    auto rules = getAllRules();
    bool recursive = allFolders;
    bool addJunk = false;
    if ( !folder )
    {
        folder = getInbox();
        addJunk = junk;
    }
    bool aOK = runRules( rules, folder, recursive );
    if ( addJunk )
        aOK = aOK && runRules( rules, getJunkFolder(), recursive );
    return aOK;
}

bool COutlookAPI::runRule( const std::shared_ptr< Outlook::Rule > &rule, std::shared_ptr< Outlook::Folder > folder, bool allFolders, bool junk )
{
    if ( !rule )
        return false;

    fIncludeJunkFolderWhenRunningOnAllFolders = junk;

    bool recursive = allFolders;
    if ( !folder )
    {
        folder = getInbox();
    }
    return runRules( { rule }, folder, recursive );
}

bool COutlookAPI::runAllRules( const std::shared_ptr< Outlook::Folder > &folder )
{
    return runRules( {}, folder );
}

bool COutlookAPI::runAllRulesOnAllFolders()
{
    auto allRules = getAllRules();
    auto inbox = getInbox();
    auto junk = getJunkFolder();

    bool retVal = true;

    int numFolders = recursiveSubFolderCount( inbox.get() );

    auto msg = QString( "Running All Rules on All Folders:" );
    auto totalFolders = numFolders + ( junk ? 1 : 0 );
    emit sigInitStatus( msg, totalFolders );

    if ( inbox )
        retVal = runRules( allRules, inbox, true, msg ) && retVal;

    if ( junk && fIncludeJunkFolderWhenRunningOnAllFolders )
        retVal = runRules( allRules, junk, false, msg ) && retVal;
    return retVal;
}

bool COutlookAPI::runAllRulesOnTrashFolder()
{
    auto allRules = getAllRules();
    auto folder = getTrashFolder();

    bool retVal = true;

    int numFolders = 1;

    auto msg = QString( "Running All Rules on Trash Folder:" );
    emit sigInitStatus( msg, numFolders );

    if ( folder )
        retVal = runRules( allRules, folder, true, msg ) && retVal;
    return retVal;
}

bool COutlookAPI::runAllRulesOnJunkFolder()
{
    auto allRules = getAllRules();
    auto folder = getJunkFolder();

    bool retVal = true;

    int numFolders = 1;

    auto msg = QString( "Running All Rules on Junk Folder:" );
    emit sigInitStatus( msg, numFolders );

    if ( folder )
        retVal = runRules( allRules, folder, true, msg ) && retVal;
    return retVal;
}

bool COutlookAPI::runRule( std::shared_ptr< Outlook::Rule > rule, const std::shared_ptr< Outlook::Folder > &folder )
{
    return runRules( std::vector< std::shared_ptr< Outlook::Rule > >( { rule } ), folder );
}

std::shared_ptr< Outlook::Rules > COutlookAPI::selectRules()
{
    if ( !selectAccount( true ) )
        return {};

    auto store = connectToException( fAccount->DeliveryStore() );
    if ( !store )
        return {};

    auto rules = store->GetRules();
    return getRules( rules );
}

std::shared_ptr< Outlook::Rules > COutlookAPI::getRules( Outlook::Rules *item )
{
    if ( !item )
        return {};
    return connectToException( std::shared_ptr< Outlook::Rules >( item ) );
}

std::shared_ptr< Outlook::Rule > COutlookAPI::findRule( const QString &rule )
{
    getRules();

    if ( !fRules )
        return {};

    for ( int ii = 1; ii <= fRules->Count(); ++ii )
    {
        auto currRule = getRule( fRules->Item( ii ) );
        if ( !currRule )
            continue;
        if ( ruleNameForRule( currRule, true ) == rule )
            return currRule;
        if ( ruleNameForRule( currRule, false ) == rule )
            return currRule;
        if ( currRule->Name() == rule )
            return currRule;
    }
    return {};
}

std::shared_ptr< Outlook::Rule > COutlookAPI::getRule( Outlook::_Rule *item )
{
    if ( !item )
        return {};
    return connectToException( std::make_shared< Outlook::Rule >( item ) );
}

std::optional< QStringList > COutlookAPI::getRecipients( Outlook::Rule *rule, QStringList *msgs )
{
    if ( !rule || !rule->Conditions() )
        return {};

    auto cond = rule->Conditions()->SenderAddress();
    if ( !cond )
    {
        if ( msgs )
            msgs->push_back( QString( "Internal error" ) );
        return {};
    }

    QStringList addresses;
    if ( cond->Enabled() )
    {
        addresses << toStringList( cond->Address() );
    }
    return addresses;
}

bool COutlookAPI::skipRule( const std::shared_ptr< Outlook::Rule > &rule ) const
{
    for ( auto &&ii : fRulesToSkip )
    {
        QRegularExpression regex( ii );
        auto ruleName = rule->Name();
        auto match = regex.match( ruleName, QRegularExpression::MatchType::PartialPreferCompleteMatch );
        bool partialMatchAllowed = ( ii.indexOf( "^" ) == -1 || ii.indexOf( "$" ) == -1 );
        if ( match.hasPartialMatch() || match.hasMatch() )
        {
            if ( ( partialMatchAllowed && match.hasPartialMatch() ) || match.hasMatch() )
                return true;
        }
    }
    return false;
}

std::vector< std::shared_ptr< Outlook::Rule > > COutlookAPI::getAllRules()
{
    getRules();
    if ( !fRules )
        return {};

    std::vector< std::shared_ptr< Outlook::Rule > > rules;
    rules.reserve( fRules->Count() );
    auto numRules = fRules->Count();
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = getRule( fRules->Item( ii ) );
        if ( skipRule( rule ) )
            continue;

        rules.push_back( rule );
    }
    return rules;
}

bool COutlookAPI::runRules( std::vector< std::shared_ptr< Outlook::Rule > > rules, std::shared_ptr< Outlook::Folder > folder, bool recursive, const std::optional< QString > &perFolderMsg /*={}*/ )
{
    if ( !folder )
        folder = rootFolder();

    if ( !folder )
        return false;

    slotClearCanceled();

    auto folderPtr = reinterpret_cast< Outlook::MAPIFolder * >( folder.get() );
    auto folderTypeID = qRegisterMetaType< Outlook::MAPIFolder * >( "MAPIFolder*", &folderPtr );

    auto msg = QString( "Running Rules on '%1':" ).arg( folderDisplayPath( folder ) );

    if ( perFolderMsg.has_value() )
    {
        emit sigIncStatusValue( perFolderMsg.value() );
    }

    if ( rules.empty() )
        rules = getAllRules();
    emit sigInitStatus( msg, static_cast< int >( rules.size() ) );

    for ( auto &&rule : rules )
    {
        if ( canceled() )
            return false;

        if ( !rule || !rule->Enabled() )
            continue;

        auto inboxPtr = fInbox.get();
        emit sigStatusMessage( QString( "Running Rule: %1 on Folder: %2" ).arg( rule->Name() ).arg( folderDisplayPath( folder ) ) );
        rule->Execute( false, QVariant( folderTypeID, &folderPtr ) );
        emit sigIncStatusValue( msg );
    }

    bool retVal = true;
    if ( recursive )
    {
        auto childFolders = getFolders( folder, false );

        for ( auto &&ii : childFolders )
        {
            retVal = runRules( rules, ii, recursive, perFolderMsg ) && retVal;
        }
    }
    sigStatusFinished( msg );
    return retVal;
}

bool COutlookAPI::addDisplayNamesToRule( Outlook::Rule *rule, const QStringList &displayNames, QStringList &msgs )
{
    if ( displayNames.isEmpty() )
        return true;

    if ( !rule || !rule->Conditions() )
        return false;

    auto header = rule->Conditions()->MessageHeader();
    if ( !header )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }

    QStringList text;
    if ( header->Enabled() )
        text = toStringList( header->Text() );

    text = getFromMessageHeaderStrings( QStringList() << text << displayNames );
    header->SetEnabled( true );
    header->SetText( text );

    return true;
}

bool COutlookAPI::addRecipientsToRule( Outlook::Rule *rule, const TEmailAddressList &recipients, QStringList &msgs )
{
    if ( recipients.empty() )
        return true;

    if ( !rule || !rule->Conditions() )
        return false;

    auto senderAddress = rule->Conditions()->SenderAddress();
    if ( !senderAddress )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }

    auto addresses = mergeRecipients( rule, recipients, &msgs );
    if ( !addresses.has_value() )
        return false;

    senderAddress->SetAddress( addresses.value() );
    senderAddress->SetEnabled( true );

    return true;
}

bool COutlookAPI::addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs )
{
    if ( recipients.isEmpty() )
        return true;

    if ( !rule || !rule->Conditions() )
        return false;

    auto senderAddress = rule->Conditions()->SenderAddress();
    if ( !senderAddress )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }

    auto addresses = mergeRecipients( rule, recipients, &msgs );
    if ( !addresses.has_value() )
        return false;

    senderAddress->SetAddress( addresses.value() );
    senderAddress->SetEnabled( true );

    return true;
}

bool COutlookAPI::addSubjectsToRule( Outlook::Rule *rule, const QStringList &subjects, QStringList &msgs )
{
    if ( subjects.isEmpty() )
        return true;

    if ( !rule || !rule->Conditions() )
        return false;

    auto header = rule->Conditions()->Subject();
    if ( !header )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }

    QStringList text;
    if ( header->Enabled() )
        text = toStringList( header->Text() );

    text = mergeStringLists( text, subjects, true );
    header->SetEnabled( true );
    header->SetText( text );

    return true;
}

bool COutlookAPI::addOutlookContactsToRule( Outlook::Rule *rule, const QStringList &outlookContacts, QStringList &msgs )
{
    return true;
}
