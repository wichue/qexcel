#ifndef QMAINWINDOW_H
#define QMAINWINDOW_H

#include <QMainWindow>
#include <QSqlDatabase>

QT_BEGIN_NAMESPACE
namespace Ui { class qMainWindow; }
QT_END_NAMESPACE

class qMainWindow : public QMainWindow
{
    Q_OBJECT

public:
    qMainWindow(QWidget *parent = nullptr);
    ~qMainWindow();

private:
    Ui::qMainWindow *ui;

    void init_db();
    //读取excel
    void readExcel(QString &path);

    void importSheet(const QVariant &rangeValue, QString excel, QString sheet);
private slots:
    void btn_import_slot();

private:
    QSqlDatabase _sqlite_db;
    const QString _db_file = "sqlite.db";
};
#endif // QMAINWINDOW_H
