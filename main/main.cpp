#include "ContactsView.h"
#include "FoldersView.h"
#include "RulesView.h"
#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    CContactsView cview;
    cview.show();

    //CRulesView rview;
    //rview.show();

    CFoldersView fview;
    fview.show();

    return a.exec();
}
