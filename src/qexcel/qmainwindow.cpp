#include "qmainwindow.h"
#include "ui_qmainwindow.h"

#include <QDebug>
#include <QFileDialog>
#include <QStandardPaths>
#include <QMessageBox>

#include <QSqlQuery>
#include <QDateTime>
#include <QAxObject>
#include <QDateTime>

#include <QSqlError>

qMainWindow::qMainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::qMainWindow)
{
    ui->setupUi(this);

    init_db();

    // 获取当前时间（包含毫秒）
    QDateTime currentDateTime = QDateTime::currentDateTime();
    // 格式化为字符串，包含毫秒
    QString timeStr = currentDateTime.toString("yyyy-MM-dd-hh-mm-ss-zzz");
    qDebug() << "当前时间：" << timeStr;  // 例如：2023-12-14 15:30:45.123

    connect(ui->btn_import,&QPushButton::clicked,this,&qMainWindow::btn_import_slot);
}

qMainWindow::~qMainWindow()
{
    _sqlite_db.close();
    delete ui;
}

void qMainWindow::init_db()
{
    //创建数据库
    _sqlite_db = QSqlDatabase::addDatabase("QSQLITE");
    _sqlite_db.setDatabaseName(_db_file);
    if(!_sqlite_db.open())
    {
        QMessageBox::warning(this,"数据库打开错误",_sqlite_db.lastError().text());
        return;
    }
}

// 导入表格
void qMainWindow::btn_import_slot()
{
    QString dbpath = QFileDialog::getOpenFileName(this,"打开",QStandardPaths::writableLocation(QStandardPaths::DesktopLocation),
                                          "xlsx(*.xlsx);;xls(*.xls);;all(*.*)");
    if(dbpath.isEmpty())
    {
        return;
    }

    readExcel(dbpath);
}

/**
 * @brief 读取excel文件
 * @param path excel文件路径
 */
void qMainWindow::readExcel(QString &path)
{
    //连接Excel控件
    QAxObject *excel;
    excel = new QAxObject("Excel.Application");
    //不显示窗体看效果
    excel->setProperty("Visible",false);
    //不显示警告看效果
    excel->setProperty("DisplayAlerts",false);


    QAxObject *workbooks;
    //按路径打开excel文件
    workbooks = excel->querySubObject("WorkBooks");

    //获取excel
    QAxObject *workbook;
    workbook = workbooks->querySubObject("Open(QString&)",path);

    //获取 excel 文件名
    QString excel_name = workbook->property("Name").toString();
    qDebug() << "文件名:" << excel_name;

    //获取所有sheet
    QAxObject *worksheets;
    worksheets = workbook->querySubObject("WorkSheets");

    //获取 Sheet 数量
    int sheetCount = worksheets->property("Count").toInt();

    qDebug() << "Sheet 数量：" << sheetCount;

    for(int index=0;index<sheetCount;index++)
    {
        QAxObject *worksheet;
        //获取第n个sheet
        worksheet = worksheets->querySubObject("Item(int)",index + 1);

        //获取 sheet 名称
        QString sheetName = worksheet->property("Name").toString();
        qDebug() << "工作表名称:" << sheetName;

        QAxObject *usedrange;
        //获取该sheet的数据范围
        usedrange = worksheet->querySubObject("UsedRange");
        //获取矩形区域的值
        QVariant rangeValue = usedrange->dynamicCall("Value");

        QVariantList varRows = rangeValue.toList();
        int row11 = varRows.size();
        QVariantList rowdata = varRows[0].toList();
        int col11 = rowdata.size();
        qDebug() << "行数：" <<row11<<"；列数："<<col11;

    #if 0 //获取行和列的另一种方式
        QAxObject *rows;
        QAxObject *columns;
        //获取行数和列数
        rows = usedrange->querySubObject("Rows");
        int iRows = rows->property("Count").toInt();
        columns = usedrange->querySubObject("Columns");
        int iColumns = columns->property("Count").toInt();

        //debug行数和列数
        qDebug() << QString("行数为:%1  列数为:%2").arg(QString::number(iRows).toInt()).arg(QString::number(iColumns).toInt());
    #endif

        importSheet(rangeValue,excel_name,sheetName);
    }

    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    if(excel)
    {
        delete excel;
        excel = nullptr;
    }
//    return rangeValue;
}

/**
 * @brief 把一个sheet导入数据库
 * @param rangeValue    sheet数据区域
 * @param excel         excel文件名
 * @param sheet         sheet名
 */
void qMainWindow::importSheet(const QVariant &rangeValue,QString excel, QString sheet)
{
    //获取所有行，每一行是一个元素
    QVariantList varRows = rangeValue.toList();

    if(varRows.isEmpty())
    {
        QMessageBox::warning(this,"错误","导入表格数据为空");
        return;
    }

    //行数
    int iRows = varRows.size();

    //获取第一行数据
    QVariantList row_head = varRows[0].toList();
    //列数
    int iColumns = row_head.size();

    //数据库表名：excel文件名+sheet表名
    QString sheet_name = excel + "" + sheet;

    QVector<QVariantList> vl;
    vl.resize(iColumns);

    //创建表sql
    QString create_sql = "create table \"";

    QDateTime currentDateTime = QDateTime::currentDateTime();
    QString timeStr = currentDateTime.toString("yyyyMMddhhmmsszzz");
    create_sql += sheet_name;
    create_sql += "\"(";

    //表头：column+index
    for(int index = 0;index < iColumns ;index++)
    {
        create_sql += "\"column" + QString::number(index) + "\"" + " varchar(255)";
        if(index < iColumns - 1)
            create_sql += ",";
    }

    create_sql += ");";

    qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<""<<create_sql;

    //给每一列赋值
    for(int i =0; i < varRows.size(); i++)
    {
        QVariantList rowData = varRows[i].toList();

        for(int j = 0;j < iColumns ;j++)
        {
            vl[j].append(rowData[j].toString());
        }
    }

    QSqlQuery query;

    //创建表
    bool ret = query.exec(create_sql);
    if(!ret)
    {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<""<<query.lastError().text();
    }

    //获取当前系统日期时间
    QDateTime dbImportTime = QDateTime::currentDateTime();
    QString str = dbImportTime.toString("yyyy.MM.dd hh:mm:ss.zzz ddd");

//    QString head = row_head[index].toString() + "column" + QString::number(index);

    //插入数据，库表的第一行是sheet的表头
    QString insert_sql = "insert into \"";
    insert_sql += sheet_name;
    insert_sql += "\"(";
    QString clms = "";
    QString values = "values(";
    for(int index = 0;index < iColumns ;index++)
    {
        clms +=   "column" + QString::number(index);
        values += ":column" + QString::number(index);
        if(index < iColumns - 1)
        {
            clms += ",";
            values += ",";
        }
    }
    values += ")";
    insert_sql = insert_sql + clms + ") " + values + ";";
    qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"insert_sql"<<insert_sql;

    //开启事务
    _sqlite_db.transaction();

    ret = query.prepare(insert_sql);
    if(!ret)
    {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<""<<query.lastError().text();
    }

    for(int index = 0;index < iColumns ;index++)
    {
        query.bindValue(":column" + QString::number(index),vl[index]);
    }

    //执行sql语句
    ret = query.execBatch();
    if(!ret)
    {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<""<<query.lastError().text()<<" "<<_sqlite_db.lastError().text();
    }
    //事务结束，执行操作
    ret = _sqlite_db.commit();
    if(!ret)
    {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<""<<query.lastError().text()<<" "<<_sqlite_db.lastError().text();
    }

    query.clear();


    //导入数据库信息显示到Qtextedit
    QString strdatetime = QDateTime::currentDateTime().toString("yyyy.MM.dd hh:mm:ss");
//    ui->textEdit->append(QString("[%1] 数据库导入成功:共计%2行,%3列.").arg(strdatetime).arg(iRows).arg(iColumns));

    QMessageBox::information(this,"成功!",QString("数据库导入成功! 共计%1行,%2列!").arg(iRows).arg(iColumns));
}





