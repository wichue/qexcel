#include "qmainwindow.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    qMainWindow w;
    w.show();
    return a.exec();
}
