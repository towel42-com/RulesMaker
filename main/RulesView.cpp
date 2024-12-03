#include "RulesView.h"
#include "RulesModel.h"
#include "ui_RulesView.h"
#include "MSOUTL.h"
#include <QTimer>

CRulesView::CRulesView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CRulesView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload( true ); } );
}

void CRulesView::init()
{
    fImpl->setupUi( this );

    fModel = std::make_shared< CRulesModel >( this );
    fImpl->rules->setModel( fModel.get() );
    connect( fImpl->rules->selectionModel(), &QItemSelectionModel::currentChanged, this, &CRulesView::slotItemSelected );
    connect(
        fModel.get(), &CRulesModel::sigFinishedLoading,
        [ = ]()
        {
            fImpl->rules->resizeColumnToContents( 0 );
            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
        } );

    connect( fModel.get(), &CRulesModel::sigSetStatus, this, &CRulesView::sigSetStatus );
    connect(
        fModel.get(), &CRulesModel::sigSetStatus,
        [ = ]( int curr, int max )
        {
            if ( ( max > 10 ) && ( curr == 1 ) || ( ( curr % 10 ) == 0 ) )
            {
                fImpl->rules->resizeColumnToContents( 0 );
            }
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

bool CRulesView::ruleSelected() const
{
    auto idx = fImpl->rules->currentIndex();
    return fModel->getRuleItem( idx ) != nullptr;
}

QString CRulesView::folderForSelectedRule() const
{
    auto rule = selectedRule();
    if ( !rule )
        return {};

    auto moveAction = rule->Actions()->MoveToFolder();
    if ( !moveAction || !moveAction->Enabled() || !moveAction->Folder() )
        return {};

    auto folderName = moveAction->Folder()->FolderPath();
    return folderName;
}

std::shared_ptr< Outlook::Rule > CRulesView::selectedRule() const
{
    auto idx = fImpl->rules->currentIndex();
    return fModel->getRule( idx );
}

void CRulesView::runSelectedRule() const
{
    auto idx = fImpl->rules->currentIndex();
    return fModel->runRule( idx );
}

void CRulesView::slotItemSelected( const QModelIndex &index )
{
    fImpl->name->clear();

    if ( index.isValid() )
    {
        auto item = fModel->getRuleItem( index );
        if ( item )
        {
            auto row = item->row();
            auto col = item->column();
            fImpl->name->setText( item->text() );
        }
    }
    emit sigRuleSelected();
}

bool CRulesView::addRule( const QString &destFolder, const QStringList &rules, QStringList &msgs )
{
    return fModel->addRule( destFolder, rules, msgs );
}

bool CRulesView::addToSelectedRule( const QStringList &rules, QStringList &msgs )
{
    auto rule = selectedRule();
    if ( !rule )
        return false;

    return fModel->addToRule( rule, rules, msgs );
}
