#include "EmailView.h"
#include "GroupedEmailModel.h"

#include "ui_EmailView.h"
#include "MSOUTL.h"

#include <list>
#include <set>
#include <QTimer>

CEmailView::CEmailView( QWidget *parent ) :
    CWidgetWithStatus( parent ),
    fImpl( new Ui::CEmailView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload( true ); } );
}

void CEmailView::init()
{
    fImpl->setupUi( this );
    setStatusLabel( "Grouping Emails:" );

    fGroupedModel = new CGroupedEmailModel( this );
    fImpl->groupedEmails->setModel( fGroupedModel );

    connect( fImpl->groupedEmails, &QTreeView::doubleClicked, this, &CEmailView::slotItemDoubleClicked );
    connect( fImpl->groupedEmails->selectionModel(), &QItemSelectionModel::selectionChanged, this, &CEmailView::slotSelectionChanged );
    connect(
        fGroupedModel, &CGroupedEmailModel::sigFinishedGrouping,
        [ = ]()
        {
            fImpl->groupedEmails->expandAll();
            fImpl->groupedEmails->resizeColumnToContents( 0 );
            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
        } );
    connect( fGroupedModel, &CGroupedEmailModel::sigSetStatus, [ = ]( int curr, int max ) { emit sigSetStatus( statusLabel(), curr, max ); } );
    connect(
        fGroupedModel, &CGroupedEmailModel::sigSetStatus,
        [ = ]( int curr, int max )
        {
            if ( ( max > 100 ) && ( curr == 1 ) || ( ( curr % 100 ) == 0 ) )
            {
                auto model = fImpl->groupedEmails->model();
                auto root = model->index( 0, 0 );
                fImpl->groupedEmails->expand( root );
                auto numRows = model->rowCount( root );
                for ( int ii = 0; ii < numRows; ++ii )
                {
                    auto idx = model->index( ii, 0, root );
                    fImpl->groupedEmails->expand( idx );
                }

                fImpl->groupedEmails->resizeColumnToContents( 0 );
            }
        } );

    setWindowTitle( QObject::tr( "Inbox Emails" ) );
}

CEmailView::~CEmailView()
{
}

void CEmailView::clear()
{
    if ( fGroupedModel )
        fGroupedModel->clear();
}

void CEmailView::clearSelection()
{
    fImpl->groupedEmails->clearSelection();
    fImpl->groupedEmails->setCurrentIndex( {} );
    slotSelectionChanged();
}

void CEmailView::slotSelectionChanged()
{
    auto rules = getRulesForSelection();
    fImpl->rule->setText( rules.join( " or " ) );
    emit sigRuleSelected();
}

QStringList CEmailView::getRulesForSelection() const
{
    auto selection = fImpl->groupedEmails->selectionModel()->selectedIndexes();

    std::set< std::list< int > > rows;
    QStringList rules;
    for ( auto &&ii : selection )
    {
        std::list< int > currRows;
        auto currIdx = ii;
        while ( currIdx.isValid() )
        {
            currRows.push_back( currIdx.row() );
            currIdx = currIdx.parent();
        }

        if ( rows.find( currRows ) != rows.end() )
            continue;
        rows.insert( currRows );

        rules << fGroupedModel->rulesForIndex( ii );
    }
    rules.removeDuplicates();
    return rules;
}

void CEmailView::slotItemDoubleClicked( const QModelIndex &idx )
{
    if ( !idx.isValid() )
        return;
    auto item = fGroupedModel->emailItemFromIndex( idx );
    if ( !item )
        return;
    item->Display();
}

void CEmailView::reload( bool notifyOnFinish )
{
    fNotifyOnFinish = notifyOnFinish;
    fGroupedModel->reload();
}
