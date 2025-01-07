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
            resizeToContentZero( fImpl->groupedEmails, EExpandMode::eExpandAll );

            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
            fImpl->summary->setText( fGroupedModel->summary() );
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

                resizeToContentZero( fImpl->groupedEmails, EExpandMode::eNoAction );
            }
        } );

    auto updateByFilter = [ = ]()
    {
        auto filterType = getFilterType();
        setFilterType( filterType );
        COutlookAPI::instance()->setEmailFilterType( filterType );
        emit sigFilterTypeChanged();
    };

    connect( fImpl->byEmailAddress, &QRadioButton::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );
    connect( fImpl->byDisplayNames, &QRadioButton::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );
    connect( fImpl->bySubjects, &QRadioButton::toggled, [ = ]( bool /*checked */ ) { updateByFilter(); } );

    setFilterType( COutlookAPI::instance()->emailFilterType() );
    slotRunningStateChanged( false );

    setWindowTitle( QObject::tr( "Inbox Emails" ) );
}

CEmailView::~CEmailView()
{
}

EFilterType CEmailView::getFilterType() const
{
    if ( fImpl->byEmailAddress->isChecked() )
        return EFilterType::eByEmailAddress;
    else if ( fImpl->byDisplayNames->isChecked() )
        return EFilterType::eByDisplayName;
    else if ( fImpl->bySubjects->isChecked() )
        return EFilterType::eBySubject;
    else
        return EFilterType::eUnknown;
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
    fImpl->emailAddresses->setText( getEmailPatternForSelection() );
    fImpl->displayNames->setText( getDisplayNamePatternForSelection() );
    fImpl->subjects->setText( getSubjectPatternForSelection() );
    emit sigEmailSelected();
}

QString CEmailView::getDisplayNamePatternForSelection() const
{
    auto text = getDisplayNamesForSelection();
    for ( auto &&curr : text )
    {
        curr = "'" + curr + "'";
    }
    return text.join( " or " );
}

QString CEmailView::getEmailPatternForSelection() const
{
    auto text = getEmailsForSelection();
    return text.join( " or " );
}

QString CEmailView::getSubjectPatternForSelection() const
{
    auto text = getSubjectsForSelection();
    for ( auto &&curr : text )
    {
        curr = "'" + curr + "'";
    }
    return text.join( " or " );
}

QStringList CEmailView::getEmailsForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&row : rows )
    {
        retVal = mergeStringLists( retVal, fGroupedModel->matchTextForIndex( row ) );
    }
    return retVal;
}

QStringList CEmailView::getDisplayNamesForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&ii : rows )
    {
        retVal << fGroupedModel->displayNamesForIndex( ii, true );
    }
    return retVal;
}

QStringList CEmailView::getSubjectsForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&ii : rows )
    {
        retVal << fGroupedModel->subjectsForIndex( ii, true );
    }
    return retVal;
}

std::pair< QStringList, EFilterType > CEmailView::getPatternsForSelection() const
{
    if ( fImpl->byEmailAddress->isChecked() )
        return { getEmailsForSelection(), EFilterType::eByEmailAddress };
    else if ( fImpl->byDisplayNames->isChecked() )
        return { getDisplayNamesForSelection(), EFilterType::eByDisplayName };
    else if ( fImpl->bySubjects->isChecked() )
        return { getSubjectsForSelection(), EFilterType::eBySubject };
    return { QStringList(), EFilterType::eUnknown};
}

bool CEmailView::selectionHasDisplayName() const
{
    return !getDisplayNamesForSelection().empty();
}

QString CEmailView::getDisplayNameForSingleSelection() const
{
    auto displayNames = getDisplayNamesForSelection();
    if ( displayNames.length() != 1 )
        return {};
    return displayNames.front();
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
    fImpl->displayNames->setEnabled( !running );
    fImpl->emailAddresses->setEnabled( !running );
    fImpl->subjects->setEnabled( !running );
    fImpl->byEmailAddress->setEnabled( !running );
    fImpl->bySubjects->setEnabled( !running );
    fImpl->byDisplayNames->setEnabled( !running );
    if ( running )
        return;

    updateEditFields();
}

void CEmailView::updateEditFields()
{
    updateEditFields( getFilterType() );
}

void CEmailView::updateEditFields( EFilterType filterType )
{
    fImpl->emailAddresses->setEnabled( filterType == EFilterType::eByEmailAddress );
    fImpl->displayNames->setEnabled( filterType == EFilterType::eByDisplayName );
    fImpl->subjects->setEnabled( filterType == EFilterType::eBySubject );
}

void CEmailView::setFilterType( EFilterType filterType )
{
    fImpl->byEmailAddress->setChecked( filterType == EFilterType::eByEmailAddress );
    fImpl->byDisplayNames->setChecked( filterType == EFilterType::eByDisplayName );
    fImpl->bySubjects->setChecked( filterType == EFilterType::eBySubject );
    updateEditFields( filterType );
}
