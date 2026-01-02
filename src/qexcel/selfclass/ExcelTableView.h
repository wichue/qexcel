#include <QTableView>
#include <QMouseEvent>
#include <QCursor>
#include <QApplication>
#include <QDebug>
#include <QHeaderView>
#include <QScrollBar>
#include <QSqlTableModel>

// 功能：鼠标悬停在下边框时，鼠标变成上下箭头，点击可拖动表格大小
class ExcelTableView : public QTableView {
    Q_OBJECT

public:
    explicit ExcelTableView(QWidget *parent = nullptr);

    void setmodel( QSqlTableModel *tablemodel);

private:
    int getContentBottom();
    void setupUI();
private slots:
    void insertRowAct_slot();
protected:
    bool eventFilter(QObject *watched, QEvent *event) override ;
    void mouseMoveEvent(QMouseEvent *event) override ;
    void mousePressEvent(QMouseEvent *event) override ;
    void mouseReleaseEvent(QMouseEvent *event) override ;
    void leaveEvent(QEvent *event) override ;
signals:
private:
    bool m_resizing = false;
    QSqlTableModel *m_tablemodel;

    // 右键菜单
    QAction *insertRowAct;
};
