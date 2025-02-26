#include "DelayDlg.h"

#include "ui_DelayDlg.h"

#include <QTimer>

CDelayDlg::CDelayDlg( const std::function< bool() > & okToEnd, QWidget *parent ) :
    QDialog( parent ),
    fOKToEnd( okToEnd ),
    fImpl( new Ui::CDelayDlg )
{
    fImpl->setupUi( this );

    fTimer = new QTimer( this );
    fTimer->setInterval( 10 );
    fTimer->setSingleShot( false );
    connect( fTimer, &QTimer::timeout, this, &CDelayDlg::slotUpdateTime );
}

CDelayDlg::~CDelayDlg()
{
}

int CDelayDlg::exec()
{
    fStartTime = QDateTime::currentDateTime();

    fTimer->start();
    return QDialog::exec();
}

void CDelayDlg::slotUpdateTime()
{
    if ( fOKToEnd() )
    {
        fTimedOut = false;
        fCancelled = false;
        QTimer::singleShot( 0, this, &CDelayDlg::accept );
        return;
    }

    auto endTime = fStartTime.addMSecs( fTimeOut );
    auto msecsRemaining = QDateTime::currentDateTime().msecsTo( endTime );

    if ( msecsRemaining < 0 )
    {
        fTimedOut = true;
        fCancelled = false;
        QTimer::singleShot( 0, this, &CDelayDlg::accept );
        return;
    }

    auto label = tr( "Time Remaining: %1.%2s" ).arg( msecsRemaining / 1000 ).arg( msecsRemaining % 1000, 3, 10, QChar( '0' ) );
    fImpl->timeLabel->setText( label );
}
