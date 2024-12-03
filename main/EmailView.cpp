#include "EmailView.h"
#include "GroupedEmailModel.h"

#include "ui_EmailView.h"
#include "MSOUTL.h"

#include <list>
#include <set>
#include <QTimer>

CEmailView::CEmailView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CEmailView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload( true ); } );
}

void CEmailView::init()
{
    fImpl->setupUi( this );

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
    connect( fGroupedModel, &CGroupedEmailModel::sigSetStatus, this, &CEmailView::sigSetStatus );

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

void CEmailView::slotSelectionChanged()
{
    auto rules = getRulesForSelection();
    if ( rules.empty() )
        return;

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

void CEmailView::setOnlyProcessUnread( bool value )
{
    fGroupedModel->setOnlyGroupUnread( value );
}

bool CEmailView::onlyProcessUnread() const
{
    return fGroupedModel->onlyGroupUnread();
}

void CEmailView::setProcessAllEmailWhenLessThan200Emails( bool value )
{
    fGroupedModel->setProcessAllEmailWhenLessThan200Emails( value );
}

bool CEmailView::processAllEmailWhenLessThan200Emails() const
{
    return fGroupedModel->processAllEmailWhenLessThan200Emails();
}
