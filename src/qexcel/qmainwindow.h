#ifndef QMAINWINDOW_H
#define QMAINWINDOW_H

#include <QMainWindow>
#include <QSqlDatabase>
#include <QSqlTableModel>
#include <QTableView>
#include <QVector>
#include <QLabel>
#include <QScrollArea>
#include <QVBoxLayout>

QT_BEGIN_NAMESPACE
namespace Ui { class qMainWindow; }
QT_END_NAMESPACE

struct DB_Table {
    QString tablename; //表名
    int row; //行数
    int column; //列数
    QSqlTableModel *tablemodel;
    QTableView *tableview;
    QAction *copyAct;

    QLabel *label;

    int filter_row; //当前过滤后的行数
    int filter_high; //当前过滤后显示区的高度,包含标签和表格
};

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

    bool query_table(DB_Table &table, QString line, int &view_high);
    void copyData(QTableView *view);

    void addTable(QString tableName, int columnCount);

private:
    void paintEvent(QPaintEvent *);
private slots:
    void btn_import_slot();
    void checkbox_changed_slot();
    void query_slot();

private:
    QScrollArea *_scroll_area;
    QVBoxLayout *_layout;

    QSqlDatabase _sqlite_db;
    QSqlTableModel *modelDspDB;
    const QString _db_file = "sqlite.db";

    QVector<DB_Table> _vTables;
};
#endif // QMAINWINDOW_H
