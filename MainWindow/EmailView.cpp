#include "EmailView.h"
#include "Models/EmailModel.h"
#include "OutlookAPI/OutlookAPI.h"

#include "ui_EmailView.h"

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

    fGroupedModel = new CEmailModel( this );
    fImpl->groupedEmails->setModel( fGroupedModel );

    connect( fImpl->groupedEmails, &QTreeView::doubleClicked, this, &CEmailView::slotItemDoubleClicked );
    connect( fImpl->groupedEmails->selectionModel(), &QItemSelectionModel::selectionChanged, this, &CEmailView::slotSelectionChanged );
    connect(
        fGroupedModel, &CEmailModel::sigFinishedGrouping,
        [ = ]()
        {
            fImpl->groupedEmails->expandAll();
            fImpl->groupedEmails->resizeColumnToContents( 0 );
            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
        } );
    connect( fGroupedModel, &CEmailModel::sigSetStatus, [ = ]( int curr, int max ) { emit sigSetStatus( statusLabel(), curr, max ); } );
    connect(
        fGroupedModel, &CEmailModel::sigSetStatus,
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

    auto updateByFilter = [ = ]( bool byEmail )
    {
        fImpl->fromNames->setEnabled( !byEmail );
        fImpl->emailAddresses->setEnabled( byEmail );
        COutlookAPI::instance()->setEmailFilterByEmail( byEmail );
    };
    connect(
        fImpl->byEmailAddress, &QRadioButton::toggled,
        [ = ]( bool checked )
        {
            updateByFilter( checked );
        } );
    connect(
        fImpl->byFromNames, &QRadioButton::toggled,
        [ = ]( bool checked )
        {
            updateByFilter( !checked );
        } );

    slotRunningStateChanged( false );

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
    fImpl->emailAddresses->setText( getDisplayTextForSelection() );
    fImpl->fromNames->setText( getFromTextForSelection() );
    emit sigEmailSelected();
}

QString CEmailView::getFromTextForSelection() const
{
    return {};
}

QString CEmailView::getDisplayTextForSelection() const
{
    auto text = getMatchTextForSelection();
    return text.join( " or " );
}

QStringList CEmailView::getMatchTextForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&row : rows )
    {
        retVal << fGroupedModel->matchTextForIndex( row );
    }
    retVal.removeDuplicates();
    retVal.removeAll( QString() );
    return retVal;
}

QString CEmailView::getEmailDisplayNameForSelection() const
{
    auto rows = getSelectedRows();
    if ( rows.empty() )
        return {};
    return fGroupedModel->displayNameForIndex( rows.front() );
}

QModelIndexList CEmailView::getSelectedRows() const
{
    auto selection = fImpl->groupedEmails->selectionModel()->selectedIndexes();

    std::set< std::list< int > > rows;
    QModelIndexList retVal;
    for ( auto &&ii : selection )
    {
        std::list< int > currRows;
        auto currIdx = ii;
        while ( currIdx.isValid() )
        {
            currRows.push_back( currIdx.row() );
            currIdx = currIdx.parent();
        }

        currRows.sort();
        if ( rows.find( currRows ) != rows.end() )
            continue;
        rows.insert( currRows );
        retVal << ii;
    }
    return retVal;
}

void CEmailView::slotItemDoubleClicked( const QModelIndex &idx )
{
    if ( !idx.isValid() )
        return;
    fGroupedModel->displayEmail( idx );
}

void CEmailView::reload( bool notifyOnFinish )
{
    fNotifyOnFinish = notifyOnFinish;
    fGroupedModel->reload();
}

void CEmailView::slotRunningStateChanged( bool running )
{
    if ( running )
        return;

    if ( COutlookAPI::instance()->emailFilterByEmail() )
        fImpl->byEmailAddress->animateClick();
    else
        fImpl->byFromNames->animateClick();
}
