#include "MainWindow.h"

#include <QApplication>

int main( int argc, char *argv[] )
{
    QApplication a( argc, argv );

    CMainWindow mw;
    mw.show();

    return a.exec();
}
