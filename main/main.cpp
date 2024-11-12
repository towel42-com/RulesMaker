#include "ContactsView.h"
#include "FoldersView.h"
#include "RulesView.h"
#include "EmailView.h"
#include "OutlookSetup.h"

#include <QApplication>
#include <QMessageLogger>
#include <QString>

#include <crtdbg.h>
QtMessageHandler originalHandler = nullptr;

void logToFile( QtMsgType type, const QMessageLogContext &context, const QString &msg )
{
    if ( type == QtMsgType::QtFatalMsg )
    {
        if ( msg == R"(ASSERT: "id < 0" in file qaxbase.cpp, line 3765)" )
            return;
    }
    //QString message = qFormatLogMessage( type, context, msg );
    //static FILE *f = fopen( "log.txt", "a" );
    //fprintf( f, "%s\n", qPrintable( message ) );
    //fflush( f );

    if ( originalHandler )
        (*originalHandler)( type, context, msg );
}

int main( int argc, char *argv[] )
{
    //_CrtSetReportMode( _CRT_WARN, _CRTDBG_MODE_DEBUG );
    //_CrtSetReportMode( _CRT_ERROR, _CRTDBG_MODE_DEBUG );
    //_CrtSetReportMode( _CRT_ASSERT, _CRTDBG_MODE_DEBUG );

    //originalHandler = qInstallMessageHandler( logToFile );
    QApplication a( argc, argv );

    COutlookSetup dlg;
    if ( dlg.exec() == QDialog::Rejected )
        return 0;

    //CContactsView cview;
    //cview.show();

    //CRulesView rview;
    //rview.show();

    CEmailView inboxView;
    inboxView.show();

    //CFoldersView fview;
    //fview.show();

    return a.exec();
}
