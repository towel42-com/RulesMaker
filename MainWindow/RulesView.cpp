#include "RulesView.h"
#include "Models/RulesModel.h"
#include "Models/ListFilterModel.h"
#include "OutlookAPI/OutlookAPI.h"
#include "OutlookAPI/ShowRule.h"

#include "ui_RulesView.h"

#include <QLineEdit>
#include <QTimer>
#include <QDebug>
#include <QCursor>
#include <QApplication>

CRulesView::CRulesView( QWidget *parent ) :
    CWidgetWithStatus( parent ),
    fImpl( new Ui::CRulesView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload( true ); } );
}

void CRulesView::init()
{
    fImpl->setupUi( this );
    setStatusLabel( "Loading Rules:" );

    fModel = new CRulesModel( this );
    fFilterModel = new CListFilterModel( this );
    fFilterModel->setOnlyFilterParent( true );
    fFilterModel->setLessThanOp(
        [ = ]( const QModelIndex &lhs, const QModelIndex &rhs )
        {
            auto lhsRule = fModel->getRule( lhs );
            auto rhsRule = fModel->getRule( rhs );
            return COutlookAPI::instance()->ruleLessThan( lhsRule, rhsRule );
        } );
    fFilterModel->setShowRowFunc(
        [ = ]( int sourceRow, const QModelIndex &sourceParent )
        {
            if ( fImpl->showDisabled->isChecked() && fImpl->showEnabled->isChecked() )
                return true;

            auto sourceIndex = fModel->index( sourceRow, 0, sourceParent );
            auto rule = fModel->getRule( sourceIndex );
            if ( !rule )
                return true;

            bool ruleEnabled = COutlookAPI::instance()->ruleEnabled( rule );
            if ( !fImpl->showDisabled->isChecked() && !ruleEnabled )
                return false;

            if ( !fImpl->showEnabled->isChecked() && ruleEnabled )
                return false;

            return true;
        } );
    fFilterModel->setSourceModel( fModel );
    fImpl->rules->setModel( fFilterModel );

    connect( fImpl->rules, &QTreeView::doubleClicked, this, &CRulesView::slotRuleDoubleClicked );

    connect( fImpl->deleteRule, &QToolButton::clicked, this, &CRulesView::slotDeleteCurrent );
    connect( fImpl->ruleEnabled, &QToolButton::clicked, this, &CRulesView::slotToggleCurrentEnable );
    connect( fImpl->rules->selectionModel(), &QItemSelectionModel::selectionChanged, this, &CRulesView::slotItemSelected );

    connect( COutlookAPI::instance().get(), &COutlookAPI::sigOptionChanged, this, &CRulesView::slotOptionsChanged );

    connect(
        fModel, &CRulesModel::sigFinishedLoading,
        [ = ]()
        {
            resizeToContentZero( fImpl->rules, EExpandMode::eNoAction );
            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
            fImpl->summary->setText( fModel->summary() );
        } );

    connect( fModel, &CRulesModel::sigSetStatus, [ = ]( int curr, int max ) { emit sigSetStatus( statusLabel(), curr, max ); } );
    connect(
        fModel, &CRulesModel::sigSetStatus,
        [ = ]( int curr, int max )
        {
            if ( ( max > 10 ) && ( curr == 1 ) || ( ( curr % 10 ) == 0 ) )
            {
                resizeToContentZero( fImpl->rules, EExpandMode::eNoAction );
            }
        } );
    connect(
        fImpl->filter, &QLineEdit::textChanged,
        [ = ]( const QString &filter )
        {
            fFilterModel->slotSetFilter( filter );

            if ( fFilterModel->rowCount() == 1 )
                fImpl->rules->expandAll();
        } );
    connect( fImpl->showEnabled, &QCheckBox::clicked, [ = ]() { fFilterModel->invalidateFilter(); } );
    connect( fImpl->showDisabled, &QCheckBox::clicked, [ = ]() { fFilterModel->invalidateFilter(); } );

    setWindowTitle( QObject::tr( "Rules" ) );
}

CRulesView::~CRulesView()
{
}

void CRulesView::reload( bool notifyOnFinished )
{
    fNotifyOnFinish = notifyOnFinished;
    fModel->reload();
}

void CRulesView::clear()
{
    if ( fModel )
        fModel->clear();
}

void CRulesView::clearSelection()
{
    fImpl->rules->clearSelection();
    fImpl->rules->setCurrentIndex( {} );
    slotItemSelected();
}

QModelIndex CRulesView::sourceIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() || ( idx.model() == fModel ) )
        return idx;
    return fFilterModel->mapToSource( idx );
}

QModelIndex CRulesView::selectedIndex() const
{
    if ( !fImpl->rules->selectionModel() )
        return {};

    auto selectedIndexes = fImpl->rules->selectionModel()->selectedIndexes();
    if ( selectedIndexes.isEmpty() )
        return {};
    auto selectedIndex = selectedIndexes.first();
    if ( !selectedIndex.isValid() )
        return selectedIndex;
    return sourceIndex( selectedIndex );
}

bool CRulesView::ruleSelected() const
{
    return fModel->ruleSelected( selectedIndex() );
}

QString CRulesView::folderForSelectedRule() const
{
    auto rule = selectedRule();
    return COutlookAPI::instance()->moveTargetFolderForRule( rule );
}

std::shared_ptr< Outlook::Rule > CRulesView::selectedRule() const
{
    return fModel->getRule( selectedIndex() );
}

std::list< EFilterType > CRulesView::filterTypesForSelectedRule() const
{
    auto rule = selectedRule();
    return COutlookAPI::instance()->filterTypesForRule( rule );
}

void CRulesView::slotRunningStateChanged( bool running )
{
    fImpl->ruleEnabled->setEnabled( !running );
    fImpl->deleteRule->setEnabled( !running );
    if ( !running )
        updateButtons( selectedIndex() );
}

void CRulesView::slotItemSelected()
{
    emit sigRuleSelected();
    updateButtons( selectedIndex() );
}

void CRulesView::slotDeleteCurrent()
{
    auto idx = selectedIndex();
    if ( !idx.isValid() )
        return;

    auto rule = fModel->getRule( idx );
    if ( !rule )
        return;
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->deleteRule( rule, false, true );
    updateButtons( rule );
    qApp->restoreOverrideCursor();
}

void CRulesView::slotToggleCurrentEnable()
{
    auto idx = selectedIndex();
    if ( !idx.isValid() )
        return;

    auto rule = fModel->getRule( idx );
    if ( !rule )
        return;

    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    bool enabled = fImpl->ruleEnabled->isChecked();
    auto status = enabled ? COutlookAPI::instance()->enableRule( rule, true ) : COutlookAPI::instance()->disableRule( rule, true );
    (void)status;
    updateButtons( rule );
    emit sigRuleSelected();
    qApp->restoreOverrideCursor();
}

void CRulesView::slotOptionsChanged()
{
    updateButtons( selectedIndex() );
}

void CRulesView::updateButtons( const QModelIndex &index )
{
    auto rule = fModel->getRule( sourceIndex( index ) );
    updateButtons( rule );
}

void CRulesView::updateButtons( const std::shared_ptr< Outlook::Rule > &rule )
{
    fImpl->deleteRule->setEnabled( rule && !COutlookAPI::instance()->disableRatherThanDeleteRules() );
    fImpl->ruleEnabled->setEnabled( rule != nullptr );
    fImpl->ruleEnabled->setChecked( rule && COutlookAPI::instance()->ruleEnabled( rule ) );
}

void CRulesView::slotRuleDoubleClicked()
{
    auto rule = fModel->getRule( selectedIndex() );
    if ( !rule )
        return;
    CShowRule dlg( rule, false, this );
    if ( dlg.exec() == QDialog::Accepted )
    {
        //fModel->reload();
    }
}
