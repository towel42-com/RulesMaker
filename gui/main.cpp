#include "MainWindow/MainWindow.h"
#include "Version.h"

#include <QApplication>

int main( int argc, char *argv[] )
{
    Q_INIT_RESOURCE( MainWindow );
    QApplication appl( argc, argv );

    NVersion::setupApplication( appl, true );

    CMainWindow mw;
    mw.show();

    return appl.exec();
}
