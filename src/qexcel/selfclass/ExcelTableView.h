#include <QTableView>
#include <QMouseEvent>
#include <QCursor>
#include <QApplication>
#include <QDebug>
#include <QHeaderView>
#include <QScrollBar>

// 功能：鼠标悬停在下边框时，鼠标变成上下箭头，点击可拖动表格大小
class ExcelTableView : public QTableView {
    Q_OBJECT

public:
    explicit ExcelTableView(QWidget *parent = nullptr)
        : QTableView(parent), m_resizing(false) {
        horizontalScrollBar()->installEventFilter(this);
        setupUI();
    }

protected:
    bool eventFilter(QObject *watched, QEvent *event) override {
           // 处理水平滚动条事件
           if (watched == horizontalScrollBar()) {
               switch (event->type()) {
               case QEvent::Enter:
                   m_mouseOnHorizontalScrollBar = true;
                   updateCursor();
                   return true;
               case QEvent::Leave:
                   m_mouseOnHorizontalScrollBar = false;
                   updateCursor();
                   return true;
               default:
                   break;
               }
           }

           return QTableView::eventFilter(watched, event);
       }

    void mouseMoveEvent(QMouseEvent *event) override {
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

    void mousePressEvent(QMouseEvent *event) override {
        qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"m_resizing"<<m_resizing;
        if (event->button() == Qt::LeftButton && !m_resizing) {
            if(cursor() == Qt::SizeVerCursor)
            {
                qDebug()<<"["<<__FILE__<<"]"<<__LINE__<<__FUNCTION__<<"tiaozheng";
                m_resizing = true;
            }
        }

        QTableView::mousePressEvent(event);
    }

    void mouseReleaseEvent(QMouseEvent *event) override {
        if (event->button() == Qt::LeftButton && m_resizing) {
            m_resizing = false;

            // 恢复应用程序光标
            QApplication::restoreOverrideCursor();
            setCursor(Qt::ArrowCursor);
        }

        QTableView::mouseReleaseEvent(event);
    }

    void leaveEvent(QEvent *event) override {
        if (!m_resizing) {
            setCursor(Qt::ArrowCursor);
        }
        QTableView::leaveEvent(event);
    }

signals:

private:
    void updateCursor() {
            if ( m_mouseOnHorizontalScrollBar) {
                // 在滚动条上时使用默认光标
                setCursor(Qt::ArrowCursor);
            } else {
                // 不在滚动条上时，检查是否在边框
            }
        }

    int getContentBottom() {
        // 获取视口区域（表格内容显示区域）
        QRect viewportRect = viewport()->rect();

        // 将视口坐标转换为widget坐标
//        QPoint viewportTopLeft = viewport()->mapTo(this, viewportRect.topLeft());
        QPoint viewportBottomRight = viewport()->mapTo(this, viewportRect.bottomRight());

        // 返回视口区域的底部Y坐标
        return viewportBottomRight.y();
    }

    void setupUI() {
        // 基本设置
        setMouseTracking(true);
        setSelectionBehavior(QAbstractItemView::SelectRows);
        setSelectionMode(QAbstractItemView::SingleSelection);

        // 允许调整行高
        verticalHeader()->setSectionsMovable(false);
        verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);
        verticalHeader()->setDefaultSectionSize(30);
    }

private:
    bool m_resizing = false;
    bool m_mouseOnHorizontalScrollBar = false;
};
