#include "MainWindow/MainWindow.h"
#include "Version.h"

#include <QApplication>
#include <objbase.h>

int main( int argc, char *argv[] )
{
    auto init = CoInitialize( nullptr );
    Q_ASSERT( init == S_OK );

    Q_INIT_RESOURCE( MainWindow );
    QApplication appl( argc, argv );

    NVersion::setupApplication( appl, true );

    CMainWindow mw;
    mw.show();

    return appl.exec();
}
