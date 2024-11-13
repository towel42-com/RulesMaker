#include "OutlookSetup.h"
#include "MainWindow.h"

#include <QApplication>

int main( int argc, char *argv[] )
{
    QApplication a( argc, argv );

    COutlookSetup dlg;
    if ( dlg.exec() == QDialog::Rejected )
        return 0;

    CMainWindow mw;
    mw.show();

    return a.exec();
}
