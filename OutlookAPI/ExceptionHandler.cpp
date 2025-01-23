#include "ExceptionHandler.h"
#include <QMessageBox>

std::shared_ptr< CExceptionHandler > CExceptionHandler::sInstance;

CExceptionHandler::CExceptionHandler( QWidget *parent, SPrivate ) :
    fParentWidget( parent )
{
}

std::shared_ptr< CExceptionHandler > CExceptionHandler::instance( QWidget *parentWidget /*= nullptr*/ )
{
    if ( !sInstance )
    {
        Q_ASSERT( parentWidget );
        sInstance = std::make_shared< CExceptionHandler >( parentWidget, SPrivate() );
    }
    else
    {
        Q_ASSERT( !parentWidget );
    }
    return sInstance;
}

std::shared_ptr< CExceptionHandler > CExceptionHandler::cliInstance()
{
    if ( !sInstance )
    {
        sInstance = std::make_shared< CExceptionHandler >( nullptr, SPrivate() );
    }
    return sInstance;
}

void CExceptionHandler::connectToException( QObject *obj )
{
    connect( obj, SIGNAL( exception( int, QString, QString, QString ) ), this, SLOT( slotHandleException( int, const QString &, const QString &, const QString & ) ) );
}

void CExceptionHandler::slotHandleException( int code, const QString &source, const QString &desc, const QString &help )
{
    if ( fIgnoreExceptions )
        return;

    if ( fParentWidget )
    {
        auto msg = QString( "%1 - %2: %3" ).arg( source ).arg( code );
        auto txt = "<br>" + desc + "</br>";
        if ( !help.isEmpty() )
            txt += "<br>" + help + "</br>";
        msg = msg.arg( txt );

        QMessageBox::critical( nullptr, "Exception Thrown", msg );
    }
    else
    {
        auto msg = QString( "%1 - %2:\n%3" ).arg( source ).arg( code );
        auto txt = desc + "\n";
        if ( !help.isEmpty() )
            txt += help + "\n";
        msg = msg.arg( txt );
        emit sigStatusMessage( msg );
        std::exit( 1 );
    }
}
