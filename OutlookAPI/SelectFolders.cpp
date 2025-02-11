#include "SelectFolders.h"
#include "OutlookAPI.h"

#include "Models/FoldersModel.h"
#include "Models/ListFilterModel.h"


#include "ui_SelectFolders.h"

#include <QSettings>

CSelectFolders::CSelectFolders( QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::CSelectFolders )
{
    init();
}

CSelectFolders::~CSelectFolders()
{
}

void CSelectFolders::init()
{
    fImpl->setupUi( this );
    fModel = new CFoldersModel( this );
    fModel->setCheckable( true );
    fFilterModel = new CListFilterModel( this );
    fFilterModel->setSourceModel( fModel );
    fImpl->folders->setModel( fFilterModel );

    connect(
        fModel, &CFoldersModel::sigFinishedLoading,
        [ = ]()
        {
            resizeToContentZero( fImpl->folders, EExpandMode::eCollapseAll );
            fImpl->folders->expandRecursively( fFilterModel->index( 0, 0 ) );
            auto inboxIndex = fFilterModel->mapFromSource( fModel->inboxIndex() );
            if ( inboxIndex.isValid() )
                fImpl->folders->expandRecursively( inboxIndex );
        } );

    connect( fModel, &CFoldersModel::sigFinishedLoadingChildren, [ = ]( QStandardItem * /*parent*/ ) { fFilterModel->sort( 0, Qt::SortOrder::AscendingOrder ); } );
    connect( fImpl->folders, &QTreeView::doubleClicked, this, &CSelectFolders::slotItemDoubleClicked );
}

void CSelectFolders::setFolders( const std::list< std::shared_ptr< Outlook::Folder > > &folders )
{
    fModel->setFolders( folders );
    for ( int ii = 0; ii < fModel->rowCount(); ++ii )
        fImpl->folders->expandRecursively( fFilterModel->index( ii, 0 ) );
}

std::list< std::shared_ptr< Outlook::Folder > > CSelectFolders::selectedFolders() const
{
    return fModel->selectedFolders();
}

void CSelectFolders::accept()
{
    QDialog::accept();
}

void CSelectFolders::slotItemDoubleClicked( const QModelIndex &idx )
{
    if ( !idx.isValid() )
        return;
    auto srcIdx = idx;
    if ( idx.model() != fModel )
        srcIdx = fFilterModel->mapToSource( idx );

    fModel->displayFolder( srcIdx );
}