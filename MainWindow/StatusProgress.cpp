#include "StatusProgress.h"
#include "ui_StatusProgress.h"

CStatusProgress::CStatusProgress( const QString & label, QWidget *parent ) :
    QWidget( parent ),
    fImpl( new Ui::CStatusProgress )
{
    fImpl->setupUi( this );
    fImpl->label->setText( label );
    fImpl->progressBar->setFormat( "(%v of %m - %p%)" );
}

CStatusProgress::~CStatusProgress()
{
}


void CStatusProgress::setRange( int min, int max )
{
    slotSetStatus( min, max );
}

void CStatusProgress::finished()
{
    hide();
    emit sigFinished();
}

void CStatusProgress::slotSetStatus( int curr, int max )
{
    fImpl->progressBar->setRange( 0, max );
    fImpl->progressBar->setValue( curr );
    if ( max && ( curr >= max ) )
    {
        if ( isVisible() )
        {
            finished();
        }
    }
    else
    {
        if ( !isVisible() )
        {
            show();
            emit sigShow();
        }
    }
    qApp->processEvents();
}

void CStatusProgress::slotIncValue()
{
    slotSetStatus( fImpl->progressBar->value() + 1, fImpl->progressBar->maximum() );
}
