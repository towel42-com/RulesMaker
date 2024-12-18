#include "RulesView.h"
#include "RulesModel.h"
#include "ListFilterModel.h"
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
    fImpl->deleteRule->setEnabled( false );
    fImpl->enableRule->setEnabled( false );
    fImpl->disableRule->setEnabled( false );

    connect( fImpl->deleteRule, &QToolButton::clicked, this, &CRulesView::slotDeleteCurrent );
    connect( fImpl->enableRule, &QToolButton::clicked, this, &CRulesView::slotEnableCurrent );
    connect( fImpl->disableRule, &QToolButton::clicked, this, &CRulesView::slotDisableCurrent );
    connect( fImpl->rules->selectionModel(), &QItemSelectionModel::currentChanged, this, &CRulesView::slotItemSelected );
    connect(
        fModel, &CRulesModel::sigFinishedLoading,
        [ = ]()
        {
            fImpl->rules->resizeColumnToContents( 0 );
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
                fImpl->rules->resizeColumnToContents( 0 );
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
    slotItemSelected( {} );
}

QModelIndex CRulesView::sourceIndex( const QModelIndex &idx ) const
{
    if ( !idx.isValid() || ( idx.model() == fModel ) )
        return idx;
    return fFilterModel->mapToSource( idx );
}

QModelIndex CRulesView::currentIndex() const
{
    auto filterIdx = fImpl->rules->currentIndex();
    if ( !filterIdx.isValid() )
        return filterIdx;
    return sourceIndex( filterIdx );
}

bool CRulesView::ruleSelected() const
{
    return fModel->getRuleItem( currentIndex() ) != nullptr;
}

QString CRulesView::folderForSelectedRule() const
{
    auto rule = selectedRule();
    return COutlookAPI::instance()->moveTargetFolderForRule( rule );
}

std::shared_ptr< Outlook::Rule > CRulesView::selectedRule() const
{
    return fModel->getRule( currentIndex() );
}

void CRulesView::slotItemSelected( const QModelIndex &index )
{
    emit sigRuleSelected();
    auto rule = fModel->getRule( sourceIndex( index ) );
    updateButtons( rule );
}

void CRulesView::updateButtons( const std::shared_ptr< Outlook::Rule > &rule )
{
    fImpl->deleteRule->setEnabled( rule && !COutlookAPI::instance()->disableRatherThanDeleteRules() );
    fImpl->enableRule->setEnabled( rule && !COutlookAPI::instance()->ruleEnabled( rule ) );
    fImpl->disableRule->setEnabled( rule && COutlookAPI::instance()->ruleEnabled( rule ) );
}

void CRulesView::slotDeleteCurrent()
{
    if ( !currentIndex().isValid() )
        return;
    auto rule = fModel->getRule( currentIndex() );
    if ( !rule )
        return;
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->deleteRule( rule );
    updateButtons( rule );
    qApp->restoreOverrideCursor();
}

void CRulesView::slotDisableCurrent()
{
    if ( !currentIndex().isValid() )
        return;
    auto rule = fModel->getRule( currentIndex() );
    if ( !rule )
        return;
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->disableRule( rule );
    updateButtons( rule );
    qApp->restoreOverrideCursor();
}

void CRulesView::slotEnableCurrent()
{
    if ( !currentIndex().isValid() )
        return;
    auto rule = fModel->getRule( currentIndex() );
    if ( !rule )
        return;
    qApp->setOverrideCursor( QCursor( Qt::WaitCursor ) );
    COutlookAPI::instance()->enableRule( rule );
    updateButtons( rule );
    qApp->restoreOverrideCursor();
}
