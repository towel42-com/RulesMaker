#include "RulesModel.h"
#include "OutlookAPI.h"

#include <QTimer>

#include "MSOUTL.h"

Q_DECLARE_METATYPE( std::shared_ptr< Outlook::Rule > );

CRulesModel::CRulesModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    connect( COutlookAPI::getInstance().get(), &COutlookAPI::sigRuleAdded, this, &CRulesModel::slotRuleAdded );
    connect( COutlookAPI::getInstance().get(), &COutlookAPI::sigRuleChanged, this, &CRulesModel::slotRuleChanged );
    connect( COutlookAPI::getInstance().get(), &COutlookAPI::sigRuleDeleted, this, &CRulesModel::slotRuleDeleted );
    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >();
    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >( "std::shared_ptr<Outlook::Rule>const&" );
}

void CRulesModel::reload()
{
    clear();
    COutlookAPI::getInstance()->slotClearCanceled();
    fRules = COutlookAPI::getInstance()->getRules();
    QTimer::singleShot( 0, [ = ]() { loadRules(); } );
}

CRulesModel::~CRulesModel()
{
}

void CRulesModel::clear()
{
    QStandardItemModel::clear();
    setHorizontalHeaderLabels( { "Name (Execution Order)", "Value" } );
    fRuleMap.clear();
    fReverseRuleMap.clear();
    fRules.reset();
}

void CRulesModel::loadRules()
{
    if ( !fRules )
    {
        emit sigFinishedLoading();
        return;
    }

    fCurrPos = 1;
    QTimer::singleShot( 0, this, &CRulesModel::slotLoadNextRule );
}

void CRulesModel::slotLoadNextRule()
{
    if ( COutlookAPI::getInstance()->canceled() )
    {
        clear();
        emit sigFinishedLoading();
        return;
    }

    auto rule = COutlookAPI::getInstance()->getRule( fRules->Item( fCurrPos ) );
    if ( rule )
    {
        loadRule( rule );
        emit sigSetStatus( fCurrPos, fRules->Count() );
    }
    fCurrPos++;
    if ( fCurrPos <= fRules->Count() )
    {
        QTimer::singleShot( 0, this, &CRulesModel::slotLoadNextRule );
    }
    else
        emit sigFinishedLoading();
}

bool CRulesModel::loadRule( std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return false;

    auto ruleItem = new QStandardItem( COutlookAPI::ruleNameForRule( rule, true ) );
    fRuleMap[ ruleItem ] = rule;
    fReverseRuleMap[ rule ] = ruleItem;

    this->appendRow( ruleItem );

    return true;
}

bool CRulesModel::hasChildren( const QModelIndex &parent ) const
{
    if ( QStandardItemModel::hasChildren( parent ) )
        return true;

    auto loaded = beenLoaded( parent );
    return !loaded;
}

bool CRulesModel::canFetchMore( const QModelIndex &parent ) const
{
    return ( parent.isValid() && hasChildren( parent ) && !beenLoaded( parent ) );
}

void CRulesModel::fetchMore( const QModelIndex &parent )
{
    if ( !parent.isValid() )
        return;

    auto ruleItem = getRuleItem( parent );
    if ( !ruleItem )
        return;
    auto rule = fRuleMap.find( ruleItem );
    if ( rule == fRuleMap.end() )
        return;

    loadRuleData( ruleItem, ( *rule ).second );
}

bool CRulesModel::beenLoaded( const QModelIndex &parent ) const
{
    if ( !parent.isValid() )
        return false;

    return beenLoaded( itemFromIndex( parent ) );
}

bool CRulesModel::beenLoaded( QStandardItem *parent ) const
{
    auto ruleItem = getRuleItem( parent );
    if ( !ruleItem )
        return false;

    auto pos = fBeenLoaded.find( ruleItem );
    return pos != fBeenLoaded.end();
}

void CRulesModel::loadRuleData( QStandardItem *ruleItem, std::shared_ptr< Outlook::Rule > rule )
{
    if ( beenLoaded( ruleItem ) )
        return;

    addAttribute( ruleItem, "Enabled", rule->Enabled() );
    addAttribute( ruleItem, "Execution Order", rule->ExecutionOrder() );
    addAttribute( ruleItem, "Is Local", rule->IsLocalRule() );
    addAttribute( ruleItem, "Rule Type", ( rule->RuleType() == Outlook::OlRuleType::olRuleReceive ) ? "Recieve" : "Send" );

    addConditions( ruleItem, rule );
    addExceptions( ruleItem, rule );
    addActions( ruleItem, rule );
    fBeenLoaded.insert( ruleItem );
}

void CRulesModel::addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule )
{
    return addConditions( parent, rule, false );
}

void CRulesModel::addExceptions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule )
{
    return addConditions( parent, rule, true );
}

void CRulesModel::addConditions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule, bool exceptions )
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

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::AccountRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Condition Type", toString( condition->ConditionType() ) );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::RuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, ruleName, "Yes" );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::ToOrFromRuleCondition *condition, bool from )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    auto recipients = COutlookAPI::getRecipientEmails( condition->Recipients(), {}, false );
    addAttribute( parent, ( from ? "From" : "To" ), recipients, " or " );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::TextRuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, ruleName, getValue( condition->Text(), " or " ) );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::CategoryRuleCondition *condition, const QString &ruleName )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, ruleName, getValue( condition->Categories(), " or " ) );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::FormNameRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Form Name", getValue( condition->FormName(), " or " ) );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::FromRssFeedRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "From RSS Feed", getValue( condition->FromRssFeed(), " or " ) );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::ImportanceRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Importance", toString( condition->Importance() ) );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::AddressRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    auto address = condition->Address();
    addAttribute( parent, "Address", getValue( condition->Address(), " or " ) );
    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::SenderInAddressListRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    auto addresses = COutlookAPI::getInstance()->getEmailAddresses( condition->AddressList(), false );
    addAttribute( parent, "Sender in Address List", addresses, " or " );

    return true;
}

bool CRulesModel::addCondition( QStandardItem *parent, Outlook::SensitivityRuleCondition *condition )
{
    if ( !condition )
        return false;

    if ( !condition->Enabled() )
        return false;

    addAttribute( parent, "Sensitivity", toString( condition->Sensitivity() ) );
    return true;
}

void CRulesModel::addActions( QStandardItem *parent, std::shared_ptr< Outlook::Rule > rule )
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

bool CRulesModel::addAction( QStandardItem *parent, Outlook::AssignToCategoryRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, "Set Categories To", getValue( action->Categories(), " and " ) );
    return true;
}

bool CRulesModel::addAction( QStandardItem *parent, Outlook::MarkAsTaskRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, "Mark as Task:", QString( "Yes - %1" ).arg( toString( action->MarkInterval() ) ) );
    return true;
}

bool CRulesModel::addAction( QStandardItem *parent, Outlook::MoveOrCopyRuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, actionName, action->Folder()->FullFolderPath() );
    return true;
}

bool CRulesModel::addAction( QStandardItem *parent, Outlook::NewItemAlertRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, "New Item Alert", action->Text() );
    return true;
}

bool CRulesModel::addAction( QStandardItem *parent, Outlook::PlaySoundRuleAction *action )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;
    addAttribute( parent, "Play Sound", '"' + action->FilePath() + '"' );
    return true;
}

bool CRulesModel::addAction( QStandardItem *parent, Outlook::RuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    addAttribute( parent, actionName, "Yes" );
    return true;
}

bool CRulesModel::addAction( QStandardItem *parent, Outlook::SendRuleAction *action, const QString &actionName )
{
    if ( !action )
        return false;
    if ( !action->Enabled() )
        return false;

    auto recipients = COutlookAPI::getRecipientEmails( action->Recipients(), {}, false );

    addAttribute( parent, actionName, recipients, " and " );
    return true;
}

void CRulesModel::addAttribute( QStandardItem *parent, const QString &label, bool value )
{
    return addAttribute( parent, label, value ? "Yes" : "No" );
}

void CRulesModel::addAttribute( QStandardItem *parent, const QString &label, int value )
{
    return addAttribute( parent, label, QString::number( value ) );
}

void CRulesModel::addAttribute( QStandardItem *parent, const QString &label, const char *value )
{
    return addAttribute( parent, label, QString( value ) );
}

void CRulesModel::addAttribute( QStandardItem *parent, const QString &label, QStringList value, const QString &separator )
{
    if ( value.size() > 1 )
    {
        for ( auto &&ii : value )
            ii = '"' + ii + '"';
    }
    auto text = value.join( separator );
    return addAttribute( parent, label, text );
}

void CRulesModel::addAttribute( QStandardItem *parent, const QString &label, const QString &value )
{
    auto keyItem = new QStandardItem( label + ":" );
    auto valueItem = new QStandardItem( value );
    parent->appendRow( { keyItem, valueItem } );
}

void CRulesModel::update()
{
    beginResetModel();
    loadRules();
    endResetModel();
}

QStandardItem *CRulesModel::getRuleItem( const QModelIndex &index ) const
{
    auto item = itemFromIndex( index );
    return getRuleItem( item );
}

QStandardItem *CRulesModel::getRuleItem( QStandardItem *item ) const
{
    if ( !item )
        return nullptr;
    if ( item->column() != 0 )
    {
        if ( item->parent() )
            item = item->parent()->child( item->row(), 0 );
        else
            item = this->item( item->row(), 0 );
    }
    while ( item->parent() )
        item = item->parent();
    return item;
}

void CRulesModel::runRule( const QModelIndex &index ) const
{
    auto item = itemFromIndex( index );
    return runRule( item );
}

void CRulesModel::runRule( QStandardItem *item ) const
{
    auto rule = getRule( item );
    if ( !rule )
        return;
    COutlookAPI::getInstance()->execute( rule );
}

std::shared_ptr< Outlook::Rule > CRulesModel::getRule( const QModelIndex &index ) const
{
    auto item = itemFromIndex( index );
    return getRule( item );
}

std::shared_ptr< Outlook::Rule > CRulesModel::getRule( QStandardItem *item ) const
{
    auto ruleItem = getRuleItem( item );
    if ( !ruleItem )
        return {};

    auto pos = fRuleMap.find( ruleItem );
    if ( pos == fRuleMap.end() )
        return {};
    return ( *pos ).second;
}

void CRulesModel::slotRuleAdded( const std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return;
    loadRule( rule );
}

void CRulesModel::slotRuleChanged( const std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return;
    updateRule( rule );
}

void CRulesModel::slotRuleDeleted( const std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return;
    auto pos = fReverseRuleMap.find( rule );
    QStandardItem *item = nullptr;
    if ( pos != fReverseRuleMap.end() )
    {
        item = ( *pos ).second;
        fReverseRuleMap.erase( pos );
    }

    auto pos2 = fRuleMap.find( item );
    if ( pos2 != fRuleMap.end() )
    {
        fRuleMap.erase( pos2 );
    }

    auto pos3 = fBeenLoaded.find( item );
    if ( pos3 != fBeenLoaded.end() )
    {
        fBeenLoaded.erase( pos3 );
    }

    auto idx = indexFromItem( item );
    removeRows( idx.row(), 1 );
    
    for ( auto && ii : fRuleMap )
    {
        auto && item = ii.first;
        auto &&rule = ii.second;
        item->setText( COutlookAPI::ruleNameForRule( rule, true ) );
    }
}

bool CRulesModel::updateRule( std::shared_ptr< Outlook::Rule > rule )
{
    QStandardItem *ruleItem = nullptr;
    auto pos = fReverseRuleMap.find( rule );
    if ( pos != fReverseRuleMap.end() )
    {
        ruleItem = ( *pos ).second;
        for ( auto jj = 0; jj < ruleItem->rowCount(); ++jj )
        {
            ruleItem->removeRow( jj );
        }
    }

    if ( !ruleItem )
        loadRule( rule );
    else
        loadRuleData( ruleItem, rule );
    return true;
}
