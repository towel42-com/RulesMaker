#include "FilterFromEmailView.h"
#include "Models/FilterFromEmailModel.h"
#include "OutlookAPI/OutlookAPI.h"

#include "ui_FilterFromEmailView.h"

#include <list>
#include <set>
#include <QTimer>

CFilterFromEmailView::CFilterFromEmailView( QWidget *parent ) :
    CWidgetWithStatus( parent ),
    fImpl( new Ui::CFilterFromEmailView )
{
    init();

    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload( true ); } );
}

void CFilterFromEmailView::init()
{
    fImpl->setupUi( this );
    setStatusLabel( "Grouping Emails:" );

    fGroupedModel = new CFilterFromEmailModel( this );
    fImpl->groupedEmails->setModel( fGroupedModel );

    connect( fImpl->groupedEmails, &QTreeView::doubleClicked, this, &CFilterFromEmailView::slotItemDoubleClicked );
    connect( fImpl->groupedEmails->selectionModel(), &QItemSelectionModel::selectionChanged, this, &CFilterFromEmailView::slotSelectionChanged );
    connect(
        fGroupedModel, &CFilterFromEmailModel::sigFinishedGrouping,
        [ = ]()
        {
            resizeToContentZero( fImpl->groupedEmails, EExpandMode::eExpandAll );

            if ( fNotifyOnFinish )
                emit sigFinishedLoading();
            fNotifyOnFinish = true;
            fImpl->summary->setText( fGroupedModel->summary() );
        } );
    connect( fGroupedModel, &CFilterFromEmailModel::sigSetStatus, [ = ]( int curr, int max ) { emit sigSetStatus( statusLabel(), curr, max ); } );
    connect(
        fGroupedModel, &CFilterFromEmailModel::sigSetStatus,
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
        updateEditFields();
        emit sigFilterTypeChanged();
    };
    initFilterTypes();

    connect( fImpl->byEmailAddress, &QCheckBox::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );
    connect( fImpl->bySenders, &QCheckBox::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );
    connect( fImpl->byDisplayNames, &QCheckBox::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );
    connect( fImpl->bySubjects, &QCheckBox::toggled, [ = ]( bool /*checked */ ) { updateByFilter(); } );

    slotRunningStateChanged( false );

    setWindowTitle( QObject::tr( "Inbox Emails" ) );
}

CFilterFromEmailView::~CFilterFromEmailView()
{
}

void CFilterFromEmailView::clear()
{
    if ( fGroupedModel )
        fGroupedModel->clear();
}

void CFilterFromEmailView::clearSelection()
{
    fImpl->groupedEmails->clearSelection();
    fImpl->groupedEmails->setCurrentIndex( {} );
    slotSelectionChanged();
}

void CFilterFromEmailView::slotSelectionChanged()
{
    fImpl->emailAddresses->setText( getEmailPatternForSelection() );
    fImpl->displayNames->setText( getDisplayNamePatternForSelection() );
    fImpl->subjects->setText( getSubjectPatternForSelection() );
    fImpl->senders->setText( getSenderPatternForSelection() );
    updateEditFields();
    emit sigEmailSelected();
}

QString CFilterFromEmailView::getDisplayNamePatternForSelection() const
{
    auto text = getDisplayNamesForSelection();
    for ( auto &&curr : text )
    {
        curr = "'" + curr + "'";
    }
    return text.join( " or " );
}

QString CFilterFromEmailView::getEmailPatternForSelection() const
{
    auto text = getEmailsForSelection();
    return text.join( " or " );
}

QString CFilterFromEmailView::getSubjectPatternForSelection() const
{
    auto text = getSubjectsForSelection();
    for ( auto &&curr : text )
    {
        curr = "'" + curr + "'";
    }
    return text.join( " or " );
}

QString CFilterFromEmailView::getSenderPatternForSelection() const
{
    auto text = getSendersForSelection();
    return text.join( " or " );
}

QStringList CFilterFromEmailView::getEmailsForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&row : rows )
    {
        retVal = mergeStringLists( retVal, fGroupedModel->matchTextForIndex( row ) );
    }
    return retVal;
}

QStringList CFilterFromEmailView::getDisplayNamesForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&ii : rows )
    {
        retVal << fGroupedModel->displayNamesForIndex( ii, true );
    }
    return retVal;
}

QStringList CFilterFromEmailView::getSubjectsForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&ii : rows )
    {
        retVal << fGroupedModel->subjectsForIndex( ii, true );
    }
    return retVal;
}

QStringList CFilterFromEmailView::getSendersForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&ii : rows )
    {
        retVal << fGroupedModel->sendersForIndex( ii, true );
    }
    return retVal;
}

bool CFilterFromEmailView::selectionHasSender() const
{
    return !getSendersForSelection().isEmpty();
}

std::list< std::pair< QStringList, EFilterType > > CFilterFromEmailView::getPatternsForSelection() const
{
    std::list< std::pair< QStringList, EFilterType > > retVal;
    if ( fImpl->byEmailAddress->isChecked() )
        retVal.emplace_back( getEmailsForSelection(), EFilterType::eByEmailAddressContains );
    if ( fImpl->byDisplayNames->isChecked() )
        retVal.emplace_back( getDisplayNamesForSelection(), EFilterType::eByDisplayName );
    if ( fImpl->bySubjects->isChecked() )
        retVal.emplace_back( getSubjectsForSelection(), EFilterType::eBySubject );
    if ( fImpl->bySenders->isChecked() )
        retVal.emplace_back( getSendersForSelection(), EFilterType::eBySender );
    return retVal;
}

bool CFilterFromEmailView::selectionHasDisplayName() const
{
    return !getDisplayNamesForSelection().empty();
}

QString CFilterFromEmailView::getDisplayNameForSingleSelection() const
{
    auto displayNames = getDisplayNamesForSelection();
    if ( displayNames.length() != 1 )
        return {};
    return displayNames.front();
}

QModelIndexList CFilterFromEmailView::getSelectedRows() const
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

void CFilterFromEmailView::slotItemDoubleClicked( const QModelIndex &idx )
{
    if ( !idx.isValid() )
        return;
    fGroupedModel->displayEmail( idx );
}

void CFilterFromEmailView::reload( bool notifyOnFinish )
{
    fNotifyOnFinish = notifyOnFinish;
    fGroupedModel->reload();
}

void CFilterFromEmailView::slotRunningStateChanged( bool running )
{
    fImpl->displayNames->setEnabled( !running );
    fImpl->emailAddresses->setEnabled( !running );
    fImpl->subjects->setEnabled( !running );
    fImpl->senders->setEnabled( !running );
    fImpl->byEmailAddress->setEnabled( !running );
    fImpl->bySubjects->setEnabled( !running );
    fImpl->byDisplayNames->setEnabled( !running );
    fImpl->bySenders->setEnabled( !running );
    if ( running )
        return;

    updateEditFields();
}

void CFilterFromEmailView::updateEditFields()
{
    fImpl->bySenders->setEnabled( selectionHasSender() );
    fImpl->emailAddresses->setEnabled( fImpl->byEmailAddress->isChecked() );
    fImpl->displayNames->setEnabled( fImpl->byDisplayNames->isChecked() );
    fImpl->subjects->setEnabled( fImpl->bySubjects->isChecked() );
    fImpl->senders->setEnabled( fImpl->bySenders->isChecked() );

    std::list< EFilterType > filterTypes;
    if ( fImpl->byEmailAddress->isChecked() )
        filterTypes.push_back( EFilterType::eByEmailAddressContains );
    if ( fImpl->byDisplayNames->isChecked() )
        filterTypes.push_back( EFilterType::eByDisplayName );
    if ( fImpl->bySubjects->isChecked() )
        filterTypes.push_back( EFilterType::eBySubject );
    if ( fImpl->bySenders->isChecked() )
        filterTypes.push_back( EFilterType::eBySender );
    COutlookAPI::instance()->setEmailFilterTypes( filterTypes );
}

void CFilterFromEmailView::initFilterTypes()
{
    auto filterTypes = COutlookAPI::instance()->emailFilterTypes();
    for ( auto &&ii : filterTypes )
    {
        switch ( ii )
        {
            case EFilterType::eByEmailAddressContains:
                fImpl->byEmailAddress->setChecked( true );
                break;
            case EFilterType::eByDisplayName:
                fImpl->byDisplayNames->setChecked( true );
                break;
            case EFilterType::eBySubject:
                fImpl->bySubjects->setChecked( true );
                break;
            case EFilterType::eBySender:
                fImpl->bySenders->setChecked( true );
                break;
            default:
                break;
        }
    }
}
