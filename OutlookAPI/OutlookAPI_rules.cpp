#include "OutlookAPI.h"

#include <QMessageBox>
#include <QStandardItem>
#include "MSOUTL.h"

static void addAttribute( QStandardItem *parent, const QString &label, const QString &value );
static void addAttribute( QStandardItem *parent, const QString &label, QStringList value, const QString &separator );
static void addAttribute( QStandardItem *parent, const QString &label, bool value );
static void addAttribute( QStandardItem *parent, const QString &label, int value );
static void addAttribute( QStandardItem *parent, const QString &label, const char *value );

static void addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule );
static void addExceptions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule );
static void addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule, bool exceptions );

static bool addCondition( QStandardItem *parent, Outlook::AccountRuleCondition *condition );
static bool addCondition( QStandardItem *parent, Outlook::RuleCondition *condition, const QString &ruleName );
static bool addCondition( QStandardItem *parent, Outlook::TextRuleCondition *condition, const QString &ruleName );
static bool addCondition( QStandardItem *parent, Outlook::CategoryRuleCondition *condition, const QString &ruleName );
static bool addCondition( QStandardItem *parent, Outlook::ToOrFromRuleCondition *condition, bool from );
static bool addCondition( QStandardItem *parent, Outlook::FormNameRuleCondition *condition );
static bool addCondition( QStandardItem *parent, Outlook::FromRssFeedRuleCondition *condition );
static bool addCondition( QStandardItem *parent, Outlook::ImportanceRuleCondition *condition );
static bool addCondition( QStandardItem *parent, Outlook::AddressRuleCondition *condition );
static bool addCondition( QStandardItem *parent, Outlook::SenderInAddressListRuleCondition *condition );
static bool addCondition( QStandardItem *parent, Outlook::SensitivityRuleCondition *condition );

static void addActions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule );
static bool addAction( QStandardItem *parent, Outlook::AssignToCategoryRuleAction *action );
static bool addAction( QStandardItem *parent, Outlook::MarkAsTaskRuleAction *action );
static bool addAction( QStandardItem *parent, Outlook::MoveOrCopyRuleAction *action, const QString &actionName );
static bool addAction( QStandardItem *parent, Outlook::NewItemAlertRuleAction *action );
static bool addAction( QStandardItem *parent, Outlook::PlaySoundRuleAction *action );
static bool addAction( QStandardItem *parent, Outlook::RuleAction *action, const QString &actionName );
static bool addAction( QStandardItem *parent, Outlook::SendRuleAction *action, const QString &actionName );


static QString conditionName( Outlook::AccountRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::RuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::TextRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::CategoryRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::ToOrFromRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::FormNameRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::FromRssFeedRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::ImportanceRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::AddressRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::SenderInAddressListRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );
static QString conditionName( Outlook::SensitivityRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly );

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

bool COutlookAPI::addRule( const std::shared_ptr< Outlook::Folder > &folder, const QStringList &rules, QStringList &msgs )
{
    if ( !folder )
        return false;

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

    if ( !addRecipientsToRule( rule.get(), rules, msgs ) )
        return false;

    auto name = ruleNameForRule( rule );
    if ( rule->Name() != name )
        rule->SetName( name );

    saveRules();

    bool retVal = runRule( rule );
    emit sigRuleAdded( rule );
    return retVal;
}

bool COutlookAPI::addToRule( std::shared_ptr< Outlook::Rule > rule, const QStringList &rules, QStringList &msgs )
{
    if ( !rule || rules.isEmpty() || !fRules )
    {
        msgs.push_back( "Parameters not set" );
        return false;
    }

    if ( !addRecipientsToRule( rule.get(), rules, msgs ) )
        return false;

    saveRules();

    bool retVal = runRule( rule );
    emit sigRuleChanged( rule );
    return retVal;
}

bool COutlookAPI::deleteRule( std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule || !fRules )
        return false;
    auto name = rule->Name();
    auto idx = rule->ExecutionOrder();
    auto ruleName = QString( "%1 (%2)" ).arg( name ).arg( idx );

    emit sigStatusMessage( QString( "Deleting Rule: %1" ).arg( ruleName ) );
    fRules->Remove( idx );

    saveRules();
    QMessageBox::information( fParentWidget, "Deleted Rule", QString( "Deleted Rule: %1" ).arg( ruleName ) );

    auto pos = fRuleBeenLoaded.find( rule );
    if ( pos != fRuleBeenLoaded.end() )
    {
        fRuleBeenLoaded.erase( pos );
    }

    emit sigRuleDeleted( rule );
    return true;
}

void COutlookAPI::loadRuleData( QStandardItem *ruleItem, std::shared_ptr< Outlook::Rule > rule )
{
    if ( ruleBeenLoaded( rule ) )
        return;

    addAttribute( ruleItem, "Name", rule->Name() );
    addAttribute( ruleItem, "Enabled", rule->Enabled() );
    addAttribute( ruleItem, "Execution Order", rule->ExecutionOrder() );
    addAttribute( ruleItem, "Is Local", rule->IsLocalRule() );
    addAttribute( ruleItem, "Rule Type", ( rule->RuleType() == Outlook::OlRuleType::olRuleReceive ) ? "Recieve" : "Send" );

    addConditions( ruleItem, rule );
    addExceptions( ruleItem, rule );
    addActions( ruleItem, rule );
    fRuleBeenLoaded.insert( rule );
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

QString COutlookAPI::ruleNameForRule( std::shared_ptr< Outlook::Rule > rule, bool forDisplay )
{
    QStringList addOns;
    if ( !rule )
        addOns << "INV-NULLPTR";

    bool isEnabled = rule ? rule->Enabled() : false;
    auto actions = rule ? rule->Actions() : nullptr;
    if ( !actions )
    {
        addOns << "INV-NOACTIONS";
    }

    Outlook::MAPIFolder *destFolder = nullptr;
    auto mvToFolderAction = actions ? actions->MoveToFolder() : nullptr;
    if ( mvToFolderAction )
    {
        isEnabled = isEnabled && mvToFolderAction->Enabled();
        destFolder = mvToFolderAction->Folder();
        if ( !destFolder )
            addOns << "NOFOLDER";
    }
    else
    {
        addOns << "INV-NOMOVEACTION";
    }

    QStringList conditions;
    if ( !forDisplay && rule && rule->Conditions() )
    {
        conditions << conditionName( rule->Conditions()->Account(), "Account", forDisplay );
        conditions << conditionName( rule->Conditions()->AnyCategory(), "AnyCategory", forDisplay );
        conditions << conditionName( rule->Conditions()->Body(), "Body", forDisplay );
        conditions << conditionName( rule->Conditions()->BodyOrSubject(), "BodyOrSubject", forDisplay );
        conditions << conditionName( rule->Conditions()->CC(), "CC", forDisplay );
        conditions << conditionName( rule->Conditions()->Category(), "Category", forDisplay );
        conditions << conditionName( rule->Conditions()->FormName(), "FormName", forDisplay );
        conditions << conditionName( rule->Conditions()->From(), "From", forDisplay );
        conditions << conditionName( rule->Conditions()->FromAnyRSSFeed(), "FromAnyRSSFeed", forDisplay );
        conditions << conditionName( rule->Conditions()->FromRssFeed(), "FromRssFeed", forDisplay );
        conditions << conditionName( rule->Conditions()->HasAttachment(), "HasAttachment", forDisplay );
        conditions << conditionName( rule->Conditions()->Importance(), "Importance", forDisplay );
        conditions << conditionName( rule->Conditions()->MeetingInviteOrUpdate(), "MeetingInviteOrUpdate", forDisplay );
        conditions << conditionName( rule->Conditions()->MessageHeader(), "MessageHeader", forDisplay );
        conditions << conditionName( rule->Conditions()->NotTo(), "NotTo", forDisplay );
        conditions << conditionName( rule->Conditions()->OnLocalMachine(), "OnLocalMachine", forDisplay );
        conditions << conditionName( rule->Conditions()->OnOtherMachine(), "OnOtherMachine", forDisplay );
        conditions << conditionName( rule->Conditions()->OnlyToMe(), "OnlyToMe", forDisplay );
        conditions << conditionName( rule->Conditions()->RecipientAddress(), "RecipientAddress", forDisplay );
        //conditions << conditionRuleName( rule->Conditions()->SenderAddress(), "SenderAddress", forDisplay );
        conditions << conditionName( rule->Conditions()->SenderInAddressList(), "SenderInAddressList", forDisplay );
        conditions << conditionName( rule->Conditions()->Sensitivity(), "Sensitivity", forDisplay );
        conditions << conditionName( rule->Conditions()->SentTo(), "SentTo", forDisplay );
        conditions << conditionName( rule->Conditions()->Subject(), "Subject", forDisplay );
        conditions << conditionName( rule->Conditions()->ToMe(), "ToMe", forDisplay );
        conditions << conditionName( rule->Conditions()->ToOrCc(), "ToOrCc", forDisplay );
    }

    if ( !isEnabled )
        conditions << ( forDisplay ? "(Disabled)" : "<Disabled>" );

    QString ruleName;
    if ( forDisplay && rule )
        ruleName = rule->Name();
    else
        ruleName = ruleNameForFolder( reinterpret_cast< Outlook::Folder * >( destFolder ) );

    if ( ruleName.isEmpty() )
        ruleName = "<UNNAMED RULE>";

    conditions.removeAll( QString() );
    conditions.sort();

    addOns.removeAll( QString() );
    addOns.sort();

    auto suffixes = QStringList() << ruleName << addOns.join( " " ) << conditions.join( " " ) << ( ( forDisplay ) ? ( rule ? QString( "(%1)" ).arg( rule->ExecutionOrder() ) : QString( "(INV_EXECUTION_ORDER)" ) ) : QString() );
    suffixes.removeAll( QString() );

    return suffixes.join( " " ).trimmed();
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

    if ( junk )
        retVal = runRules( allRules, junk, false, msg ) && retVal;
    return retVal;
}

bool COutlookAPI::runRule( std::shared_ptr< Outlook::Rule > rule, const std::shared_ptr< Outlook::Folder > &folder )
{
    return runRules( std::vector< std::shared_ptr< Outlook::Rule > >( { rule } ), folder );
}

bool COutlookAPI::enableAllRules()
{
    if ( !fRules )
        return false;

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
        rule->SetEnabled( true );
        numChanged++;
    }
    if ( canceled() )
        return false;

    if ( numChanged != 0 )
        saveRules();

    QMessageBox::information( fParentWidget, R"(Enable All Rules)", QString( "%1 rules enabled" ).arg( numChanged ) );

    return numChanged != 0;
}

bool COutlookAPI::mergeRules()
{
    if ( !fRules )
        return false;

    auto numRules = fRules->Count();
    emit sigInitStatus( "Merging Rules:", numRules );
    std::map< QString, std::shared_ptr< Outlook::Rule > > rules;
    std::list< int > toRemove;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        emit sigIncStatusValue( "Merging Rules:" );
        if ( canceled() )
            return false;

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule || !rule->Enabled() )
            continue;

        auto from = rule->Conditions()->SenderAddress();
        if ( !from || !from->Enabled() )
            continue;

        auto moveAction = rule->Actions()->MoveToFolder();
        if ( !moveAction || !moveAction->Enabled() )
            continue;

        auto key = moveAction->Folder()->FullFolderPath();
        auto pos = rules.find( key );
        if ( pos == rules.end() )
        {
            rules[ key ] = rule;
        }
        else
        {
            auto mergedRecipients = mergeRecipients( ( *pos ).second.get(), rule.get(), nullptr );
            if ( !mergedRecipients.has_value() )
                continue;

            rule->SetEnabled( false );
            ( *pos ).second->Conditions()->SenderAddress()->SetAddress( mergedRecipients.value() );
            toRemove.push_front( ii );
        }
    }
    if ( canceled() )
        return false;
    auto numChanged = toRemove.size();
    for ( auto &&ii : toRemove )
    {
        if ( canceled() )
            return false;
        fRules->Remove( ii );
    }

    if ( !toRemove.empty() )
        saveRules();

    QMessageBox::information( fParentWidget, R"(Merge Rules by Target Folder)", QString( "%1 rules deleted" ).arg( numChanged ) );

    return !toRemove.empty();
}

bool COutlookAPI::moveFromToAddress()
{
    if ( !fRules )
        return false;

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
    QMessageBox::information( fParentWidget, R"(Move "From" to "Address")", QString( "%1 rules modified" ).arg( numChanged ) );
    return numChanged;
}

bool COutlookAPI::renameRules()
{
    if ( !fRules )
        return false;

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
        QMessageBox::information( fParentWidget, "Renamed Rules", QString( "No rules needed renaming" ) );
        return 0;
    }
    QStringList tmp;
    for ( auto &&ii : changes )
    {
        tmp << "<li>" + ii.first->Name() + " => " + ii.second + "</li>";
    }
    auto msg = QString( "Rules to be changed:<ul>%1</ul>Continue?" ).arg( tmp.join( "\n" ) );
    auto process = QMessageBox::information( fParentWidget, "Renamed Rules", msg, QMessageBox::Yes | QMessageBox::No );
    if ( process == QMessageBox::No )
        return 0;

    emit sigInitStatus( "Renaming Rules:", static_cast< int >( changes.size() ) );
    for ( auto &&ii : changes )
    {
        ii.first->SetName( ii.second );
        emit sigIncStatusValue( "Renaming Rules:" );
    }
    saveRules();

    return changes.size();
}

bool COutlookAPI::sortRules()
{
    if ( !fRules )
        return false;

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
    bool changed = false;
    auto pos = 1;
    emit sigInitStatus( "Recomputing Execution Order:", numRules );

    for ( auto &&ii : rules )
    {
        if ( canceled() )
            return false;

        changed = changed || ( ii->ExecutionOrder() != pos );
        ii->SetExecutionOrder( pos++ );
        emit sigIncStatusValue( "Recomputing Execution Order:" );
    }
    if ( changed )
        saveRules();
    return changed;
}

void COutlookAPI::saveRules()
{
    emit sigStatusMessage( QString( "Saving Rules" ) );
    fRules->Save( true );
}

std::shared_ptr< Outlook::Rules > COutlookAPI::selectRules()
{
    if ( !fAccount )
    {
        if ( !selectAccount( true ) )
            return {};
    }

    if ( !fAccount || fAccount->isNull() )
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
        auto variant = cond->Address();
        if ( variant.type() == QVariant::Type::String )
            addresses << variant.toString();
        else if ( variant.type() == QVariant::Type::StringList )
            addresses << variant.toStringList();
    }
    return addresses;
}

std::vector< std::shared_ptr< Outlook::Rule > > COutlookAPI::getAllRules()
{
    if ( !fRules )
        return {};

    std::vector< std::shared_ptr< Outlook::Rule > > rules;
    rules.reserve( fRules->Count() );
    auto numRules = fRules->Count();
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = getRule( fRules->Item( ii ) );
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

bool COutlookAPI::addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs )
{
    if ( recipients.isEmpty() )
        return true;

    if ( !rule || !rule->Conditions() )
        return false;

    auto cond = rule->Conditions()->SenderAddress();
    if ( !cond )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }

    auto addresses = mergeRecipients( rule, recipients, &msgs );
    if ( !addresses.has_value() )
        return false;

    cond->SetAddress( addresses.value() );
    cond->SetEnabled( true );

    return true;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs )
{
    auto lhsRecipients = getRecipients( lhs, msgs );
    if ( !lhsRecipients.has_value() )
        return {};
    if ( !lhsRecipients )
        return rhs;
    lhsRecipients.value() << rhs;
    lhsRecipients.value().removeDuplicates();
    return lhsRecipients;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs )
{
    auto lhsRecipients = getRecipients( lhs, msgs );
    auto rhsRecipients = getRecipients( rhs, msgs );
    if ( !lhsRecipients.has_value() && !rhsRecipients.has_value() )
        return {};
    if ( lhsRecipients && !rhsRecipients )
        return lhsRecipients;
    if ( !lhsRecipients && rhsRecipients )
        return rhsRecipients;
    lhsRecipients.value() << rhsRecipients.value();
    lhsRecipients.value().removeDuplicates();
    return lhsRecipients;
}


template< typename T >
static QString conditionRuleNameBase( T *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( condition && condition->Enabled() )
    {
        if ( forDisplayOnly )
            return "<" + conditionStr + ">";
        else
            return "(" + conditionStr + ")";
    }
    return {};
}

QString conditionName( Outlook::SensitivityRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->Sensitivity() );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::SenderInAddressListRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto addresses = COutlookAPI::instance()->getEmailAddresses( condition->AddressList(), false );
    auto retVal = conditionStr + "=";
    retVal += addresses.join( " or " );

    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::AddressRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->Address(), " or " );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::ImportanceRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->Importance() );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::FromRssFeedRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->FromRssFeed(), " or " );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::FormNameRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->FormName(), " or " );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::ToOrFromRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=";

    auto recipients = COutlookAPI::getEmailAddresses( condition->Recipients(), {}, false );
    retVal += recipients.join( " or " );

    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::CategoryRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->Categories(), " or " );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::TextRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->Text(), " or " );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::RuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=Yes";
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

QString conditionName( Outlook::AccountRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->ConditionType() );
    return conditionRuleNameBase( condition, retVal, forDisplayOnly );
}

void addExceptions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule )
{
    return addConditions( parent, rule, true );
}

void addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule, bool exceptions )
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

    found = addCondition( folder, conditions->Account() ) || found;
    found = addCondition( folder, conditions->AnyCategory(), "Any Category" ) || found;
    found = addCondition( folder, conditions->Body(), "Body" ) || found;
    found = addCondition( folder, conditions->BodyOrSubject(), "Body or Subject" ) || found;
    found = addCondition( folder, conditions->CC(), "CC" ) || found;
    found = addCondition( folder, conditions->Category(), "Category" ) || found;
    found = addCondition( folder, conditions->FormName() ) || found;
    found = addCondition( folder, conditions->From(), true ) || found;
    found = addCondition( folder, conditions->FromAnyRSSFeed(), "From Any RSS Feed" ) || found;
    found = addCondition( folder, conditions->FromRssFeed() ) || found;
    found = addCondition( folder, conditions->HasAttachment(), "Has Attachment" ) || found;
    found = addCondition( folder, conditions->Importance() ) || found;
    found = addCondition( folder, conditions->MeetingInviteOrUpdate(), "Meeting Invite Or Update" ) || found;
    found = addCondition( folder, conditions->MessageHeader(), "Message Header" ) || found;
    found = addCondition( folder, conditions->NotTo(), "Not To" ) || found;
    found = addCondition( folder, conditions->OnLocalMachine(), "On Local Machine" ) || found;
    found = addCondition( folder, conditions->OnOtherMachine(), "On Other Machine" ) || found;
    found = addCondition( folder, conditions->OnlyToMe(), "Only to Me" ) || found;
    found = addCondition( folder, conditions->RecipientAddress() ) || found;
    found = addCondition( folder, conditions->SenderAddress() ) || found;
    found = addCondition( folder, conditions->SenderInAddressList() ) || found;
    found = addCondition( folder, conditions->Sensitivity() ) || found;
    found = addCondition( folder, conditions->SentTo(), "Sent To" ) || found;
    found = addCondition( folder, conditions->Subject(), "Subject" ) || found;
    found = addCondition( folder, conditions->ToMe(), "To Me" ) || found;
    found = addCondition( folder, conditions->ToOrCc(), "To or CC" ) || found;

    if ( found )
        parent->appendRow( folder );
    else
        delete folder;
}

void addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule )
{
    return addConditions( parent, rule, false );
}

bool addCondition( QStandardItem *parent, Outlook::AccountRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Condition Type", toString( condition->ConditionType() ) );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::RuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, ruleName, "Yes" );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::ToOrFromRuleCondition *condition, bool from )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    auto recipients = COutlookAPI::getEmailAddresses( condition->Recipients(), {}, false );
    addAttribute( parent, ( from ? "From" : "To" ), recipients, " or " );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::TextRuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, ruleName, toString( condition->Text(), " or " ) );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::CategoryRuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, ruleName, toString( condition->Categories(), " or " ) );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::FormNameRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Form Name", toString( condition->FormName(), " or " ) );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::FromRssFeedRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "From RSS Feed", toString( condition->FromRssFeed(), " or " ) );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::ImportanceRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Importance", toString( condition->Importance() ) );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::AddressRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Address", toString( condition->Address(), " or " ) );
    return true;
}

bool addCondition( QStandardItem *parent, Outlook::SenderInAddressListRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    auto addresses = COutlookAPI::instance()->getEmailAddresses( condition->AddressList(), false );
    addAttribute( parent, "Sender in Address List", addresses, " or " );

    return true;
}

bool addCondition( QStandardItem *parent, Outlook::SensitivityRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Sensitivity", toString( condition->Sensitivity() ) );
    return true;
}

void addActions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return;
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

    found = addAction( folder, actions->AssignToCategory() ) || found;
    found = addAction( folder, actions->MarkAsTask() ) || found;
    found = addAction( folder, actions->CopyToFolder(), "Copy to Folder" ) || found;
    found = addAction( folder, actions->MoveToFolder(), "Move to Folder" ) || found;
    found = addAction( folder, actions->NewItemAlert() ) || found;
    found = addAction( folder, actions->PlaySound() ) || found;
    found = addAction( folder, actions->ClearCategories(), "Clear Categories" ) || found;
    found = addAction( folder, actions->Delete(), "Delete" ) || found;
    found = addAction( folder, actions->DeletePermanently(), "Delete Permanently" ) || found;
    found = addAction( folder, actions->DesktopAlert(), "Desktop Alert" ) || found;
    found = addAction( folder, actions->NotifyDelivery(), "Notify Delivery" ) || found;
    found = addAction( folder, actions->NotifyRead(), "Notify Read" ) || found;
    found = addAction( folder, actions->Stop(), "Stop" ) || found;
    found = addAction( folder, actions->CC(), "Send as CC" ) || found;
    found = addAction( folder, actions->Forward(), "Forward" ) || found;
    found = addAction( folder, actions->ForwardAsAttachment(), "Forward as Attachment" ) || found;
    found = addAction( folder, actions->Redirect(), "Redirect" ) || found;

    if ( found )
        parent->appendRow( folder );
    else
        delete folder;
}

bool addAction( QStandardItem *parent, Outlook::AssignToCategoryRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, "Set Categories To", toString( action->Categories(), " and " ) );
    return true;
}

bool addAction( QStandardItem *parent, Outlook::MarkAsTaskRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, "Mark as Task:", QString( "Yes - %1" ).arg( toString( action->MarkInterval() ) ) );
    return true;
}

bool addAction( QStandardItem *parent, Outlook::MoveOrCopyRuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, actionName, action->Folder()->FullFolderPath() );
    return true;
}

bool addAction( QStandardItem *parent, Outlook::NewItemAlertRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, "New Item Alert", action->Text() );
    return true;
}

bool addAction( QStandardItem *parent, Outlook::PlaySoundRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;
    addAttribute( parent, "Play Sound", '"' + action->FilePath() + '"' );
    return true;
}

bool addAction( QStandardItem *parent, Outlook::RuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, actionName, "Yes" );
    return true;
}

bool addAction( QStandardItem *parent, Outlook::SendRuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    auto recipients = COutlookAPI::getEmailAddresses( action->Recipients(), {}, false );

    addAttribute( parent, actionName, recipients, " and " );
    return true;
}

void addAttribute( QStandardItem *parent, const QString &label, bool value )
{
    return addAttribute( parent, label, value ? "Yes" : "No" );
}

void addAttribute( QStandardItem *parent, const QString &label, int value )
{
    return addAttribute( parent, label, QString::number( value ) );
}

void addAttribute( QStandardItem *parent, const QString &label, const char *value )
{
    return addAttribute( parent, label, QString( value ) );
}

void addAttribute( QStandardItem *parent, const QString &label, QStringList value, const QString &separator )
{
    if ( value.size() > 1 )
    {
        for ( auto &&ii : value )
            ii = '"' + ii + '"';
    }
    auto text = value.join( separator );
    return addAttribute( parent, label, text );
}

void addAttribute( QStandardItem *parent, const QString &label, const QString &value )
{
    auto keyItem = new QStandardItem( label + ":" );
    auto valueItem = new QStandardItem( value );
    parent->appendRow( { keyItem, valueItem } );
}
