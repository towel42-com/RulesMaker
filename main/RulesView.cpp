#include "RulesView.h"
#include "RulesModel.h"
#include "ui_RulesView.h"

#include <QTimer>

CRulesView::CRulesView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CRulesView )
{
    init();
    
    if ( !parent )
        QTimer::singleShot( 0, [ = ]() { reload(); } );

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
            fImpl->rules->expandAll();
            fImpl->rules->resizeColumnToContents( 0 );
            emit sigFinishedLoading();
        } );

    setWindowTitle( QObject::tr( "Rules" ) );
}

CRulesView::~CRulesView()
{
}

void CRulesView::reload()
{
    fModel->reload();
}

void CRulesView::clear()
{
    if ( fModel )
        fModel->clear();
}

void CRulesView::slotItemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;

    fImpl->name->clear();
    auto item = fModel->itemFromIndex( index );
    if ( !item )
        return;
    while ( item->parent() )
        item = item->parent();

    fImpl->name->setText( item->text() );
}
