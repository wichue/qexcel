#include "ExcelTableView.h"
#include <QAction>

ExcelTableView::ExcelTableView(QWidget *parent)
    : QTableView(parent), m_resizing(false)
{
    horizontalScrollBar()->installEventFilter(this);
    setupUI();

    insertRowAct = new QAction("插入行",this);
    addAction(insertRowAct);
    connect(insertRowAct,&QAction::triggered,this,&ExcelTableView::insertRowAct_slot);
}

void ExcelTableView::setmodel(QSqlTableModel *tablemodel)
{
    m_tablemodel = tablemodel;
}

void ExcelTableView::setupUI()
{
    // 基本设置
    setMouseTracking(true);
//    setSelectionBehavior(QAbstractItemView::SelectRows);// 选中整行
    setSelectionMode(QAbstractItemView::SingleSelection);

    // 允许调整行高
    verticalHeader()->setSectionsMovable(false);
    verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    verticalHeader()->setDefaultSectionSize(30);
}

void ExcelTableView::insertRowAct_slot()
{
    int row =0;
    int column =0;
    QModelIndexList indexes = selectionModel()->selectedIndexes();
    //没有选中任何区域时，在最后一行下面插入行
    if(indexes.size() == 0)
    {
        row = m_tablemodel->rowCount();
        column = m_tablemodel->columnCount();
    }
    else
    {
        row = indexes.at(0).row();
        column = indexes.at(0).column();
    }

    m_tablemodel->insertRow(row);
//    m_tablemodel->submitAll();

    qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"row"<<row<<"column"<<column;
}

bool ExcelTableView::eventFilter(QObject *watched, QEvent *event)
{
   // 处理水平滚动条事件
   if (watched == horizontalScrollBar()) {
       switch (event->type()) {
       case QEvent::Enter:
           setCursor(Qt::ArrowCursor);
           return true;
       case QEvent::Leave:
           return true;
       default:
           break;
       }
   }

   return QTableView::eventFilter(watched, event);
}

void ExcelTableView::mouseMoveEvent(QMouseEvent *event) {
    if (!m_resizing) {
        // 只在非调整状态下检测鼠标位置
        QPoint pos = mapFromGlobal(event->globalPos());

        // 计算内容区域的底部边界（排除滚动条）
        int contentBottom = getContentBottom();

        // 检查鼠标是否在内容区域的下边框
        bool nearBottomBorder = /*(pos.y() >= contentBottom - 5) &&*/
                               (pos.y() == contentBottom);
//            qDebug() << "pos.y():" << pos.y() << "contentBottom:" <<contentBottom;

        if (nearBottomBorder) {
            setCursor(Qt::SizeVerCursor);
        } else {
            setCursor(Qt::ArrowCursor);
        }
    } else {
        // 正在调整高度
        QPoint pos = mapFromGlobal(event->globalPos());
//            QSize newSize = size();
//            newSize.setHeight(pos.y());

//            setMaximumHeight(QWIDGETSIZE_MAX);
//            setMinimumHeight(30);
//            resize(newSize);
        //设置高度会重新布局
        if(pos.y() > 30)
            setFixedHeight(pos.y());
    }

    QTableView::mouseMoveEvent(event);
}

void ExcelTableView::mousePressEvent(QMouseEvent *event)
{
//    qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"m_resizing"<<m_resizing;
    if (event->button() == Qt::LeftButton && !m_resizing) {
        if(cursor() == Qt::SizeVerCursor)
        {
//            qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"tiaozheng";
            m_resizing = true;
        }
    }

    QTableView::mousePressEvent(event);
}

void ExcelTableView::mouseReleaseEvent(QMouseEvent *event)
{
    if (event->button() == Qt::LeftButton && m_resizing) {
        m_resizing = false;

        // 恢复应用程序光标
        QApplication::restoreOverrideCursor();
        setCursor(Qt::ArrowCursor);
    }

    QTableView::mouseReleaseEvent(event);
}

void ExcelTableView::leaveEvent(QEvent *event) {
    if (!m_resizing) {
        setCursor(Qt::ArrowCursor);
    }
    QTableView::leaveEvent(event);
}

int ExcelTableView::getContentBottom()
{
    // 获取视口区域（表格内容显示区域）
    QRect viewportRect = viewport()->rect();

    // 将视口坐标转换为widget坐标
//        QPoint viewportTopLeft = viewport()->mapTo(this, viewportRect.topLeft());
    QPoint viewportBottomRight = viewport()->mapTo(this, viewportRect.bottomRight());

    // 返回视口区域的底部Y坐标
    return viewportBottomRight.y();
}


