#include "StatusProgress.h"
#include <QLabel>
#include <QHBoxLayout>
#include <QProgressBar>
#include <QApplication>

CStatusProgress::CStatusProgress( const QString &label, QWidget *parent ) :
    QWidget( parent )
{
    setObjectName( "CStatusProgress" );
    resize( 226, 21 );
    QSizePolicy sizePolicy( QSizePolicy::Preferred, QSizePolicy::Minimum );
    sizePolicy.setHorizontalStretch( 0 );
    sizePolicy.setVerticalStretch( 0 );
    sizePolicy.setHeightForWidth( this->sizePolicy().hasHeightForWidth() );
    setSizePolicy( sizePolicy );

    auto horizontalLayout = new QHBoxLayout( this );
    horizontalLayout->setObjectName( QString::fromUtf8( "horizontalLayout" ) );
    horizontalLayout->setContentsMargins( 0, 0, 0, 0 );
    fLabel = new QLabel( this );
    fLabel->setObjectName( QString::fromUtf8( "label" ) );

    horizontalLayout->addWidget( fLabel );

    fProgressBar = new QProgressBar( this );
    fProgressBar->setObjectName( QString::fromUtf8( "fProgressBar" ) );
    fProgressBar->setValue( 24 );

    horizontalLayout->addWidget( fProgressBar );

    fLabel->setText( label );
    fProgressBar->setFormat( "(%v of %m - %p%)" );
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
    fProgressBar->setRange( 0, max );
    fProgressBar->setValue( curr );
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
    slotSetStatus( fProgressBar->value() + 1, fProgressBar->maximum() );
}
