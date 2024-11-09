#include "RulesView.h"
#include "RulesModel.h"
#include "ui_RulesView.h"

#include <QTimer>

CRulesView::CRulesView( QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CRulesView )
{
    fImpl->setupUi( this );
    connect( fImpl->addButton, &QPushButton::clicked, this, &CRulesView::addEntry );
    connect( fImpl->changeButton, &QPushButton::clicked, this, &CRulesView::changeEntry );

    QTimer::singleShot(
        0,
        [ = ]()
        {
            fModel = std::make_shared< CRulesModel >( this );
            fImpl->rules->setModel( fModel.get() );
            connect( fImpl->rules->selectionModel(), &QItemSelectionModel::currentChanged, this, &CRulesView::itemSelected );
        } );

    setWindowTitle( QObject::tr( "Rules" ) );
}

CRulesView::~CRulesView()
{
}

void CRulesView::updateOutlook()
{
    fModel->update();
}

void CRulesView::addEntry()
{
    if ( !fImpl->name->text().isEmpty() )
    {
        fModel->addItem( fImpl->name->text() );
    }

    fImpl->name->clear();
}

void CRulesView::changeEntry()
{
    QModelIndex current = fImpl->rules->currentIndex();

    if ( current.isValid() )
        fModel->changeItem( current, fImpl->name->text() );
}

void CRulesView::itemSelected( const QModelIndex &index )
{
    if ( !index.isValid() )
        return;

    QAbstractItemModel *model = fImpl->rules->model();
    fImpl->name->setText( model->data( model->index( index.row(), 0 ) ).toString() );
}
