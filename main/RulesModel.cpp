#include "RulesModel.h"
#include "OutlookHelpers.h"

#include <QTimer>
#include <QProgressDialog>
#include <QProgressBar>

#include "MSOUTL.h"

CRulesModel::CRulesModel( QObject *parent ) :
    QStandardItemModel( parent )
{
}

void CRulesModel::reload()
{
    fRules = COutlookHelpers::getInstance()->getRules( dynamic_cast< QWidget * >( parent() ) );
    if ( !fRules )
        return;

    QTimer::singleShot( 0, [ = ]() { loadRules(); } );
}

CRulesModel::~CRulesModel()
{
}

void CRulesModel::loadRules()
{
    clear();
    setHorizontalHeaderLabels( { "Name", "Value" } );
    setColumnCount( 2 );

    if ( !fRules )
        return;

    auto numRules = fRules->Count();
    QProgressDialog dlg( dynamic_cast< QWidget * >( parent() ) );
    auto bar = new QProgressBar;
    bar->setFormat( "(%v of %m - %p%)" );
    dlg.setBar( bar );
    dlg.setMinimum( 0 );
    dlg.setMaximum( numRules );
    dlg.setLabelText( "Loading Rules" );
    dlg.setMinimumDuration( 0 );
    dlg.setWindowModality( Qt::WindowModal );

    for ( int ii = 1; ii <= numRules; ++ii )
    {
        dlg.setValue( ii );
        if ( dlg.wasCanceled() )
        {
            clear();
            break;
        }

        if ( ii == 7 )
            int xyz = 0;
        auto rule = std::make_shared< Outlook::Rule >( fRules->Item( ii ) );
        if ( !rule )
            continue;

        auto ruleItem = new QStandardItem( rule->Name() );
        this->appendRow( ruleItem );

        addAttribute( ruleItem, "Enabled", rule->Enabled() );
        addAttribute( ruleItem, "Execution Order", rule->ExecutionOrder() );
        addAttribute( ruleItem, "Is Local", rule->IsLocalRule() );
        addAttribute( ruleItem, "Rule Type", ( rule->RuleType() == Outlook::OlRuleType::olRuleReceive ) ? "Recieve" : "Send" );

        addConditions( ruleItem, rule );
        addExceptions( ruleItem, rule );
        addActions( ruleItem, rule );
    }

    emit sigFinishedLoading();
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

    auto recipients = COutlookHelpers::getRecipientEmails( condition->Recipients(), {} );
    addAttribute( parent, ( from ? "From" : "To" ), recipients, " or " );
    return true;
}

QString getValue( const QVariant &variant, const QString &joinSeparator )
{
    QString retVal;
    if ( variant.type() == QVariant::Type::String )
        retVal = variant.toString();
    else if ( variant.type() == QVariant::Type::StringList )
        retVal = variant.toStringList().join( joinSeparator );
    else
        int xyz = 0;
    return retVal;
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

    auto addresses = COutlookHelpers::getInstance()->getEmailAddresses( condition->AddressList() );
    addAttribute( parent, "Sender in Address List", addresses, " or "  );

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

    auto recipients = COutlookHelpers::getRecipientEmails( action->Recipients(), {} );

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

void CRulesModel::changeItem( const QModelIndex & /*index*/, const QString & /*folderName*/ )
{
    if ( !fRules )
        return;

    //Outlook::Folder item( fItems->Item( index.row() + 1 ) );

    //item.SetName( folderName );
    ////item.Save();

    //fCache.take( index );
}

void CRulesModel::addItem( const QString & /*folderName*/ )
{
    //Outlook::Folder item( COutlookHelpers::getInstance()->outlook()->CreateItem( Outlook::OlItemType::olContactItem ) );
    //if ( !item.isNull() )
    //{
    //    item.SetName( folderName );
    //    item.Save();
    //}
}

void CRulesModel::update()
{
    beginResetModel();
    loadRules();
    endResetModel();
}
