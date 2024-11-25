#include "MainWindow.h"

#include <QApplication>

int main( int argc, char *argv[] )
{
    QApplication appl( argc, argv );

    appl.setApplicationName( "Outlook Rules Maker" );
    appl.setOrganizationDomain( "towel42.com" );
    appl.setOrganizationName( "Towel 42 Development" );
    appl.setApplicationVersion( "0.9" );

    CMainWindow mw;
    mw.show();

    return appl.exec();
}
