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

    fGroupedModel = new CEmailModel( this );
    fImpl->groupedEmails->setModel( fGroupedModel );

    connect( fImpl->groupedEmails, &QTreeView::doubleClicked, this, &CFilterFromEmailView::slotItemDoubleClicked );
    connect( fImpl->groupedEmails->selectionModel(), &QItemSelectionModel::selectionChanged, this, &CFilterFromEmailView::slotSelectionChanged );
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
        updateEditFields();
        emit sigFilterTypeChanged();
    };
    initFilterTypes();

    connect( fImpl->byEmailAddress, &QCheckBox::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );
    connect( fImpl->byDisplayNames, &QCheckBox::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );
    connect( fImpl->bySubjects, &QCheckBox::toggled, [ = ]( bool /*checked */ ) { updateByFilter(); } );
    connect( fImpl->byOutlookContacts, &QCheckBox::toggled, [ = ]( bool /*checked*/ ) { updateByFilter(); } );

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
    fImpl->outlookContacts->setText( getOutlookContactsPatternForSelection() );
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

QString CFilterFromEmailView::getOutlookContactsPatternForSelection() const
{
    auto text = getOutlookContactsForSelection();
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

QStringList CFilterFromEmailView::getOutlookContactsForSelection() const
{
    auto rows = getSelectedRows();
    QStringList retVal;
    for ( auto &&ii : rows )
    {
        retVal << fGroupedModel->outlookContactsForIndex( ii, true );
    }
    return retVal;
}

bool CFilterFromEmailView::selectionHasOutlookContact() const
{
    return !getOutlookContactsForSelection().isEmpty();
}

std::list< std::pair< QStringList, EFilterType > > CFilterFromEmailView::getPatternsForSelection() const
{
    std::list< std::pair< QStringList, EFilterType > > retVal;
    if ( fImpl->byEmailAddress->isChecked() )
        retVal.emplace_back( getEmailsForSelection(), EFilterType::eByEmailAddress );
    if ( fImpl->byDisplayNames->isChecked() )
        retVal.emplace_back( getDisplayNamesForSelection(), EFilterType::eByDisplayName );
    if ( fImpl->bySubjects->isChecked() )
        retVal.emplace_back( getSubjectsForSelection(), EFilterType::eBySubject );
    if ( fImpl->byOutlookContacts->isChecked() )
        retVal.emplace_back( getOutlookContactsForSelection(), EFilterType::eByOutlookContact );
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
    fImpl->outlookContacts->setEnabled( !running );
    fImpl->byEmailAddress->setEnabled( !running );
    fImpl->bySubjects->setEnabled( !running );
    fImpl->byDisplayNames->setEnabled( !running );
    fImpl->byOutlookContacts->setEnabled( !running );
    if ( running )
        return;

    updateEditFields();
}

void CFilterFromEmailView::updateEditFields()
{
    fImpl->byOutlookContacts->setEnabled( selectionHasOutlookContact() );
    fImpl->emailAddresses->setEnabled( fImpl->byEmailAddress->isChecked() );
    fImpl->displayNames->setEnabled( fImpl->byDisplayNames->isChecked() );
    fImpl->subjects->setEnabled( fImpl->bySubjects->isChecked() );
    fImpl->outlookContacts->setEnabled( fImpl->byOutlookContacts->isChecked() );

    std::list< EFilterType > filterTypes;
    if ( fImpl->byEmailAddress->isChecked() )
        filterTypes.push_back( EFilterType::eByEmailAddress );
    if ( fImpl->byDisplayNames->isChecked() )
        filterTypes.push_back( EFilterType::eByDisplayName );
    if ( fImpl->bySubjects->isChecked() )
        filterTypes.push_back( EFilterType::eBySubject );
    if ( fImpl->byOutlookContacts->isChecked() )
        filterTypes.push_back( EFilterType::eByOutlookContact );
    COutlookAPI::instance()->setEmailFilterTypes( filterTypes );
}

void CFilterFromEmailView::initFilterTypes()
{
    auto filterTypes = COutlookAPI::instance()->emailFilterTypes();
    for ( auto &&ii : filterTypes )
    {
        switch ( ii )
        {
            case EFilterType::eByEmailAddress:
                fImpl->byEmailAddress->setChecked( true );
                break;
            case EFilterType::eByDisplayName:
                fImpl->byDisplayNames->setChecked( true );
                break;
            case EFilterType::eBySubject:
                fImpl->bySubjects->setChecked( true );
                break;
            case EFilterType::eByOutlookContact:
                fImpl->byOutlookContacts->setChecked( true );
                break;
            default:
                break;
        }
    }
}
