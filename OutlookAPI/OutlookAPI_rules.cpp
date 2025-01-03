#include "OutlookAPI.h"

#include <QMessageBox>
#include <QStandardItem>
#include <QDebug>
#include <QRegularExpression>

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
static bool addCondition( QStandardItem *parent, Outlook::ToOrFromRuleCondition *condition, bool from );   // from or sentTo
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

bool COutlookAPI::addRule( const std::shared_ptr< Outlook::Folder > &folder, const QStringList &rules, EFilterType patternType, QStringList &msgs )
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

    return addToRule( rule, rules, patternType, msgs );
}

bool COutlookAPI::addToRule( std::shared_ptr< Outlook::Rule > rule, const QStringList &rules, EFilterType patternType, QStringList &msgs )
{
    if ( !rule || rules.isEmpty() || !fRules )
    {
        msgs.push_back( "Parameters not set" );
        return false;
    }

    switch ( patternType )
    {
        case EFilterType::eByEmailAddress:
            {
                if ( !addRecipientsToRule( rule.get(), rules, msgs ) )
                    return false;
            }
            break;
        case EFilterType::eByDisplayName:
            {
                if ( !addDisplayNamesToRule( rule.get(), rules, msgs ) )
                    return false;
            }
            break;
        case EFilterType::eBySubject:
            {
                if ( !addSubjectsToRule( rule.get(), rules, msgs ) )
                    return false;
            }
        default:
            return false;
    }

    auto name = ruleNameForRule( rule, false );
    if ( rule->Name() != name )
        rule->SetName( name );

    saveRules();

    bool retVal = runRule( rule );
    emit sigRuleChanged( rule );
    return retVal;
}

bool COutlookAPI::ruleEnabled( const std::shared_ptr< Outlook::Rule > &rule )
{
    if ( !rule )
        return false;
    return rule->Enabled();
}

bool COutlookAPI::deleteRule( std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule || !fRules )
        return false;
    if ( disableRatherThanDeleteRules() )
        return disableRule( rule );

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

bool COutlookAPI::disableRule( const std::shared_ptr< Outlook::Rule > &rule )
{
    auto name = rule->Name();
    auto idx = rule->ExecutionOrder();
    auto ruleName = QString( "%1 (%2)" ).arg( name ).arg( idx );

    emit sigStatusMessage( QString( "Disabling Rule: %1" ).arg( ruleName ) );
    rule->SetEnabled( false );

    saveRules();
    QMessageBox::information( fParentWidget, "Disabled Rule", QString( "Disabled Rule: %1" ).arg( ruleName ) );

    emit sigRuleChanged( rule );
    return true;
}

bool COutlookAPI::enableRule( const std::shared_ptr< Outlook::Rule > &rule )
{
    auto name = rule->Name();
    auto idx = rule->ExecutionOrder();
    auto ruleName = QString( "%1 (%2)" ).arg( name ).arg( idx );

    emit sigStatusMessage( QString( "Disabling Rule: %1" ).arg( ruleName ) );
    rule->SetEnabled( true );

    saveRules();
    QMessageBox::information( fParentWidget, "Enabled Rule", QString( "Enabled Rule: %1" ).arg( ruleName ) );

    emit sigRuleChanged( rule );
    return true;
}

void COutlookAPI::loadRuleData( QStandardItem *ruleItem, std::shared_ptr< Outlook::Rule > rule, bool force )
{
    if ( ruleBeenLoaded( rule ) )
    {
        if ( !force )
            return;
        else
            ruleItem->removeRows( 0, ruleItem->rowCount() );
    }

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

EFilterType COutlookAPI::filterTypeForRule( const std::shared_ptr< Outlook::Rule > &rule ) const
{
    if ( !rule )
        return {};

    auto conditions = rule->Conditions();
    if ( !conditions )
        return {};
    auto senderAddress = conditions->SenderAddress();
    if ( senderAddress && senderAddress->Enabled() )
        return EFilterType::eByEmailAddress;

    auto header = rule->Conditions()->MessageHeader();
    if ( header && header->Enabled() )
        return EFilterType::eByDisplayName;

    auto subject = rule->Conditions()->Subject();
    if ( subject && subject->Enabled() )
        return EFilterType::eBySubject;

    return EFilterType::eUnknown;
}

QString COutlookAPI::ruleNameForRule( std::shared_ptr< Outlook::Rule > rule, bool forDisplay, bool rawName )
{
    QStringList addOns;
    if ( !rule )
    {
        if ( rawName )
            return {};
        addOns << "INV-NULLPTR";
    }
    if ( rawName )
        return rule->Name();

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
        conditions << conditionName( rule->Conditions()->Account(), "Account", false );
        conditions << conditionName( rule->Conditions()->AnyCategory(), "AnyCategory", false );
        conditions << conditionName( rule->Conditions()->Body(), "Body", false );
        conditions << conditionName( rule->Conditions()->BodyOrSubject(), "BodyOrSubject", false );
        conditions << conditionName( rule->Conditions()->CC(), "CC", false );
        conditions << conditionName( rule->Conditions()->Category(), "Category", false );
        conditions << conditionName( rule->Conditions()->FormName(), "FormName", false );
        conditions << conditionName( rule->Conditions()->From(), "From", false );
        conditions << conditionName( rule->Conditions()->FromAnyRSSFeed(), "FromAnyRSSFeed", false );
        conditions << conditionName( rule->Conditions()->FromRssFeed(), "FromRssFeed", false );
        conditions << conditionName( rule->Conditions()->HasAttachment(), "HasAttachment", false );
        conditions << conditionName( rule->Conditions()->Importance(), "Importance", false );
        conditions << conditionName( rule->Conditions()->MeetingInviteOrUpdate(), "MeetingInviteOrUpdate", false );
        conditions << conditionName( rule->Conditions()->MessageHeader(), "MessageHeader", false );
        conditions << conditionName( rule->Conditions()->NotTo(), "NotTo", false );
        conditions << conditionName( rule->Conditions()->OnLocalMachine(), "OnLocalMachine", false );
        conditions << conditionName( rule->Conditions()->OnOtherMachine(), "OnOtherMachine", false );
        conditions << conditionName( rule->Conditions()->OnlyToMe(), "OnlyToMe", false );
        conditions << conditionName( rule->Conditions()->RecipientAddress(), "RecipientAddress", false );
        //conditions << conditionName( rule->Conditions()->SenderAddress(), "SenderAddress", false );
        conditions << conditionName( rule->Conditions()->SenderInAddressList(), "SenderInAddressList", false );
        conditions << conditionName( rule->Conditions()->Sensitivity(), "Sensitivity", false );
        conditions << conditionName( rule->Conditions()->SentTo(), "SentTo", false );
        conditions << conditionName( rule->Conditions()->Subject(), "Subject", false );
        conditions << conditionName( rule->Conditions()->ToMe(), "ToMe", false );
        conditions << conditionName( rule->Conditions()->ToOrCc(), "ToOrCc", false );
    }

    if ( !isEnabled )
    {
        if ( !forDisplay )
            conditions << "<Disabled>";
    }

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

    text = mergeStringLists( text, displayNames, true );
    for ( auto &&ii : text )
    {
        ii = "From: " + ii;
    }
    header->SetEnabled( true );
    header->SetText( text );

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