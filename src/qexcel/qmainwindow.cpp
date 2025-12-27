#include "qmainwindow.h"
#include "ui_qmainwindow.h"

#include <QDebug>
#include <QFileDialog>
#include <QStandardPaths>
#include <QMessageBox>
#include <QHeaderView>
#include <QSqlQuery>
#include <QDateTime>
#include <QAxObject>
#include <QDateTime>

#include <QSqlError>
#include <QLineEdit>
#include <QSqlRecord>
#include <QClipboard>

#define COLUMN_HIGH 33      // 行高
#define LABEL_HIGH  33      // 标签的高度
#define MAX_SHOW_ROW 10     // 最大显示行数
#define VIEW_START_HIGH 50  // 显示区域的起始高度

qMainWindow::qMainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::qMainWindow)
{
    ui->setupUi(this);
    this->setWindowTitle("qexcel");

    ui->btn_import->move(this->width() - ui->btn_import->width() - 10,10);
    ui->checkBox_edit->move(this->width() - ui->btn_import->width() - ui->checkBox_edit->width() - 10,10);
    ui->lineEdit_Query->resize(this->width() - ui->btn_import->width() - ui->checkBox_edit->width() - 30,ui->lineEdit_Query->height());

    _scroll_area = new QScrollArea(this);
    _layout = new QVBoxLayout(ui->widget);
    _scroll_area->move(0,50);

    init_db();

    _scroll_area->setWidget(ui->widget);
    _scroll_area->setWidgetResizable(true);

    // 获取当前时间（包含毫秒）
    QDateTime currentDateTime = QDateTime::currentDateTime();
    // 格式化为字符串，包含毫秒
    QString timeStr = currentDateTime.toString("yyyy-MM-dd-hh-mm-ss-zzz");
    qDebug() << "当前时间：" << timeStr;  // 例如：2023-12-14 15:30:45.123

    connect(ui->btn_import,&QPushButton::clicked,this,&qMainWindow::btn_import_slot);
    connect(ui->lineEdit_Query,&QLineEdit::textChanged,this,&qMainWindow::query_slot);
    connect(ui->checkBox_edit,&QCheckBox::stateChanged,this,&qMainWindow::checkbox_changed_slot);
}

qMainWindow::~qMainWindow()
{
    _sqlite_db.close();
    delete ui;
}

void qMainWindow::init_db()
{
    //打开数据库
    _sqlite_db = QSqlDatabase::addDatabase("QSQLITE");
    _sqlite_db.setDatabaseName(_db_file);
    if(!_sqlite_db.open())
    {
        QMessageBox::warning(this,"数据库打开错误",_sqlite_db.lastError().text());
        return;
    }

    QSqlQuery query(_sqlite_db);
    if (query.exec("SELECT name FROM sqlite_master WHERE type='table'")) {
        while (query.next()) {
            QString tableName = query.value(0).toString();

            // 获取列数
            QSqlRecord record = _sqlite_db.record(tableName);
            int columnCount = record.count();

            addTable(tableName,columnCount);
        }
    } else {
        qDebug() << "查询失败:" << query.lastError();
    }
}

void qMainWindow::addTable(QString tableName, int columnCount)
{
    DB_Table table;
    table.tablename = tableName;
    table.column = columnCount;

    table.label = new QLabel;
    table.label->hide();
    table.label->resize(300,LABEL_HIGH);

    table.tableview = new QTableView;
    table.tableview->horizontalHeader()->setVisible(false);
    table.tableview->verticalHeader()->setDefaultSectionSize(COLUMN_HIGH);
    table.tableview->hide();
    table.tableview->setEditTriggers(QAbstractItemView::NoEditTriggers);
    table.tableview->setSelectionMode(QAbstractItemView::ContiguousSelection);
    table.tableview->setContextMenuPolicy(Qt::ActionsContextMenu);

    table.copyAct = new QAction("复制",this);
    table.tableview->addAction(table.copyAct);
    connect(table.copyAct,&QAction::triggered,
            [=]()
                {
                    copyData(table.tableview);
                }
            );
    table.copyAct->setShortcut(tr("Ctrl+C"));

    _layout->addWidget(table.label);
    _layout->addWidget(table.tableview);

    table.tablemodel = new QSqlTableModel;
    table.tableview->setModel(table.tablemodel);
    table.tablemodel->setTable(tableName);
    table.tablemodel->select();
    while(table.tablemodel->canFetchMore())
    {
        table.tablemodel->fetchMore();
    }
    table.row = table.tablemodel->rowCount();
    _vTables.append(table);

    qDebug() << QString("表: %1,行数：%2,列数: %3").arg(tableName).arg(table.tablemodel->rowCount()).arg(columnCount);
}

void qMainWindow::checkbox_changed_slot()
{
    switch(ui->checkBox_edit->checkState())
    {
    case Qt::Unchecked:
        for(int index=0;index<_vTables.size();index++)
        {
            _vTables[index].tableview->setEditTriggers(QAbstractItemView::NoEditTriggers);
        }
        break;
    case Qt::Checked:
    case Qt::PartiallyChecked:
        for(int index=0;index<_vTables.size();index++)
        {
            _vTables[index].tableview->setEditTriggers(QAbstractItemView::DoubleClicked);
        }
        break;

    default:
        break;
    }
}

void qMainWindow::query_slot()
{
    QString line = ui->lineEdit_Query->text().trimmed();
    if(line.isEmpty())
    {
        for(int index=0;index<_vTables.size();index++)
        {
            _vTables[index].tableview->hide();
            _vTables[index].label->hide();
        }

       return;
    }

    int view_high = VIEW_START_HIGH;
    auto it = _vTables.begin();
    while (it != _vTables.end()) {
        it->label->hide();
        it->tableview->hide();

        if (!query_table(*it,line,view_high)) {
            it = _vTables.erase(it);
        } else {
            ++it;
        }
    }
}

bool qMainWindow::query_table(DB_Table &table, QString line, int& view_high)
{
    QString cond = "";
    for(int index=0;index<table.column;index++)
    {
        cond += QString("column%1 like '%%2%'").arg(QString::number(index)).arg(line);

        if(index < table.column - 1)
            cond += " or ";
    }
    cond += "or rowid = 1";
    qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"tablename:" << table.tablename <<",cond:"<<cond;

    table.tableview->setModel(table.tablemodel);
    table.tablemodel->setFilter(cond);
    if(!table.tablemodel->select())
    {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"select error:"<<table.tablemodel->lastError().text();
        return false;
    }
    table.filter_row = table.tablemodel->rowCount();

    int show_row = 0;
    if(table.filter_row >= MAX_SHOW_ROW)
    {
        show_row = MAX_SHOW_ROW;
    }
    else
    {
        show_row = table.filter_row;
    }

//    table.filter_high = LABEL_HIGH + show_row * COLUMN_HIGH + 20; // 20冗余，滚动条会占用高度
    table.filter_high = LABEL_HIGH + 20;
    if(show_row > 1)
    {
        table.label->setText(table.tablename);
        table.label->move(10,view_high);
        table.label->show();

        table.tableview->resizeColumnsToContents();
        table.tableview->resizeRowsToContents();

        for(int index=0;index<show_row;index++)
        {
            table.filter_high += table.tableview->rowHeight(index);
        }

        table.tableview->show();
        table.tableview->move(10,view_high + LABEL_HIGH);
//        table.tableview->resize(this->width() - 20,table.filter_high - LABEL_HIGH);
        table.tableview->setFixedHeight(table.filter_high - LABEL_HIGH);

        view_high += table.filter_high + 20;
    }

    //在底部添加拉伸因子，使控件靠上
    _layout->addStretch();

    return true;
}

void qMainWindow::paintEvent(QPaintEvent *)
{
    _scroll_area->resize(this->width(),this->height() - 50);
    ui->btn_import->move(this->width() - ui->btn_import->width() - 10,10);
    ui->checkBox_edit->move(this->width() - ui->btn_import->width() - ui->checkBox_edit->width() - 10,10);
    ui->lineEdit_Query->resize(this->width() - ui->btn_import->width() - ui->checkBox_edit->width() - 30,ui->lineEdit_Query->height());
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
//    QFileInfo fileInfo(excel_fullname);
//    QString excel_name = fileInfo.completeBaseName();
//    qDebug() << "cleanName:" << excel_name;

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
}

/**
 * @brief 把一个sheet导入数据库，表名可以包含中文，但不能包含特殊字符.
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

    QString excel_real = excel;
    QString sheet_real = sheet;

    excel = excel.remove(".");
    sheet = sheet.remove(".");

    //行数
    int iRows = varRows.size();

    //获取第一行数据
    QVariantList row_head = varRows[0].toList();
    //列数
    int iColumns = row_head.size();

    //数据库表名：excel文件名+sheet表名
    QString sheet_name = excel + "" + sheet;

    QSqlQuery query;
    QString sql_query = QString(
            "SELECT count(*) FROM sqlite_master "
            "WHERE type='table' AND name='%1'"
        ).arg(sheet_name);

    if (!query.exec(sql_query)) {
        qDebug() << "Query failed:" << query.lastError().text();
        QMessageBox::warning(this,"错误",QString("导入失败 : excel : %1 ,sheet : %2, error : %3").arg(excel_real).arg(sheet_real).arg(query.lastError().text()));
        return;
    }

    if (query.next()) {
        if(query.value(0).toInt() > 0) {
            QMessageBox::warning(this,"失败",QString("存在同名表 : excel : %1 ,sheet : %2").arg(excel_real).arg(sheet_real));
            return;
        }
    }

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

    //创建表
    bool ret = query.exec(create_sql);
    if(!ret)
    {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<""<<query.lastError().text();
        QMessageBox::warning(this,"失败",QString("导入失败 : excel : %1 ,sheet : %2, error : %3").arg(excel_real).arg(sheet_real).arg(query.lastError().text()));
        query.clear();

        return;
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
        QMessageBox::warning(this,"失败",QString("导入失败 : excel : %1 ,sheet : %2, error : %3").arg(excel_real).arg(sheet_real).arg(query.lastError().text()));
        query.clear();

        return;
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
        QMessageBox::warning(this,"失败",QString("导入失败 : excel : %1 ,sheet : %2, error : %3").arg(excel_real).arg(sheet_real).arg(query.lastError().text()));
        query.clear();

        return;
    }
    //事务结束，执行操作
    ret = _sqlite_db.commit();
    if(!ret)
    {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<""<<query.lastError().text()<<" "<<_sqlite_db.lastError().text();
        QMessageBox::warning(this,"失败",QString("导入失败 : excel : %1 ,sheet : %2, error : %3").arg(excel).arg(sheet).arg(query.lastError().text()));
        query.clear();

        return;
    }

    query.clear();

    addTable(sheet_name,iColumns);

    //导入成功
    QMessageBox::information(this,"成功",QString("导入成功 : excel : %1 ,sheet : %2, 行：%3，列：%4").arg(excel_real).arg(sheet_real).arg(iRows).arg(iColumns));

    query_slot();
}

void qMainWindow::copyData(QTableView *view)
{
    int minRow =0;
    int minColumn =0;
    int maxRow =0;
    int maxColumn =0;
    QMap<QString,QString> map;
    QModelIndexList indexes = view->selectionModel()->selectedIndexes();

    QModelIndex index;
    int k = 0;

    foreach(index,indexes)
    {
        int col = index.column();
        int row = index.row();
        if(k == 0)
        {
            minColumn = col;
            minRow = row;
        }
        if(col > maxColumn)
            maxColumn = col;
        if(row > maxRow)
            maxRow = row;

        QString text = index.model()->data(index,Qt::EditRole).toString();
        //qDebug() << text;
        //qDebug() << "我是分割线";
        //QString::number,把数字转换为字符串
        map[QString::number(row) + "," +QString::number(col)] = text;
        //qDebug() << map;
        k++;
    }
    QString rs = "";
    for(int row = minRow; row<=maxRow; row++)
    {
        for(int col = minColumn; col<=maxColumn; col++)
        {
            if(col != minColumn)
                rs += "\t";
            rs+=map[QString::number(row)+","+QString::number(col)];
        }
        rs+="\r\n";
    }
    QClipboard *board=QApplication::clipboard();
    board->setText(rs);
}





