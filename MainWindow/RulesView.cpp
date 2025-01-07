#include "RulesView.h"
#include "ShowRule.h"
#include "Models/RulesModel.h"
#include "Models/ListFilterModel.h"
#include "OutlookAPI/OutlookAPI.h"

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
    fFilterModel->setSourceModel( fModel );
    fImpl->rules->setModel( fFilterModel );

    connect( fImpl->rules, &QTreeView::doubleClicked, this, &CRulesView::slotRuleDoubleClicked );

    connect( fImpl->deleteRule, &QToolButton::clicked, this, &CRulesView::slotDeleteCurrent );
    connect( fImpl->enableRule, &QToolButton::clicked, this, &CRulesView::slotEnableCurrent );
    connect( fImpl->disableRule, &QToolButton::clicked, this, &CRulesView::slotDisableCurrent );
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

EFilterType CRulesView::filterTypeForSelectedRule() const
{
    auto rule = selectedRule();
    return COutlookAPI::instance()->filterTypeForRule( rule );
}

void CRulesView::slotRunningStateChanged( bool running )
{
    fImpl->enableRule->setEnabled( !running );
    fImpl->disableRule->setEnabled( !running );
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
    if ( !selectedIndex().isValid() )
        return;
    auto rule = fModel->getRule( selectedIndex() );
    if ( !rule )
        return;
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->deleteRule( rule );
    updateButtons( rule );
    qApp->restoreOverrideCursor();
}

void CRulesView::slotDisableCurrent()
{
    if ( !selectedIndex().isValid() )
        return;
    auto rule = fModel->getRule( selectedIndex() );
    if ( !rule )
        return;
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->disableRule( rule );
    updateButtons( rule );
    qApp->restoreOverrideCursor();
}

void CRulesView::slotEnableCurrent()
{
    if ( !selectedIndex().isValid() )
        return;
    auto rule = fModel->getRule( selectedIndex() );
    if ( !rule )
        return;
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->enableRule( rule );
    updateButtons( rule );
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
    fImpl->enableRule->setEnabled( rule && !COutlookAPI::instance()->ruleEnabled( rule ) );
    fImpl->disableRule->setEnabled( rule && COutlookAPI::instance()->ruleEnabled( rule ) );
}

void CRulesView::slotRuleDoubleClicked()
{
    auto rule = fModel->getRule( selectedIndex() );
    if ( !rule )
        return;
    CShowRule dlg( rule, false, this );
    if ( dlg.exec() == QDialog::Accepted )
    {
        //COutlookAPI::instance()->saveRules();
        //fModel->reload();
    }
}
