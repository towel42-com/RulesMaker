#include "RulesModel.h"
#include "OutlookAPI.h"

#include <QTimer>

Q_DECLARE_METATYPE( std::shared_ptr< Outlook::Rule > );

CRulesModel::CRulesModel( QObject *parent ) :
    QStandardItemModel( parent )
{
    connect( COutlookAPI::instance().get(), &COutlookAPI::sigRuleAdded, this, &CRulesModel::slotRuleAdded );
    connect( COutlookAPI::instance().get(), &COutlookAPI::sigRuleChanged, this, &CRulesModel::slotRuleChanged );
    connect( COutlookAPI::instance().get(), &COutlookAPI::sigRuleDeleted, this, &CRulesModel::slotRuleDeleted );
    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >();
    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >( "std::shared_ptr<Outlook::Rule>const&" );
}

void CRulesModel::reload()
{
    clear();
    COutlookAPI::instance()->slotClearCanceled();
    fRules = COutlookAPI::instance()->getRules();
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
    fRules.first.reset();
    fRules.second = 0;
}

void CRulesModel::loadRules()
{
    if ( !fRules.first )
    {
        emit sigFinishedLoading();
        return;
    }

    fCurrPos = 1;
    QTimer::singleShot( 0, this, &CRulesModel::slotLoadNextRule );
}

void CRulesModel::slotLoadNextRule()
{
    if ( COutlookAPI::instance()->canceled() )
    {
        clear();
        emit sigFinishedLoading();
        return;
    }

    auto rule = COutlookAPI::instance()->getRule( fRules.first, fCurrPos );
    if ( rule )
    {
        loadRule( rule );
        emit sigSetStatus( fCurrPos, fRules.second );
    }
    fCurrPos++;
    if ( fCurrPos <= fRules.second )
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

    auto rule = getRule( parent );
    if ( !rule )
        return false;

    auto loaded = COutlookAPI::instance()->ruleBeenLoaded( rule );
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

    COutlookAPI::instance()->loadRuleData( ruleItem, ( *rule ).second );
}

bool CRulesModel::beenLoaded( const QModelIndex &parent ) const
{
    if ( !parent.isValid() )
        return false;

    auto rule = getRule( parent );
    if ( !rule )
        return false;
    return COutlookAPI::instance()->ruleBeenLoaded( rule );
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

void CRulesModel::updateAllRules()
{
    for ( auto &&ii : fRuleMap )
    {
        auto &&item = ii.first;
        auto &&rule = ii.second;
        item->setText( COutlookAPI::ruleNameForRule( rule, true ) );
    }
}

void CRulesModel::slotRuleAdded( const std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return;
    loadRule( rule );
    updateAllRules();
}

void CRulesModel::slotRuleChanged( const std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule )
        return;
    updateRule( rule );
    updateAllRules();
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

    auto idx = indexFromItem( item );
    removeRows( idx.row(), 1 );

    updateAllRules();
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
        COutlookAPI::instance()->loadRuleData( ruleItem, rule );
    return true;
}
