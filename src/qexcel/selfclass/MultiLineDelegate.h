#ifndef MULTILINEDELEGATE_H
#define MULTILINEDELEGATE_H

#include <QStyledItemDelegate>
#include <QPlainTextEdit>

// 自定义委托：将默认单行编辑器替换为多行编辑器
class MultiLineDelegate : public QStyledItemDelegate
{
    Q_OBJECT
public:
    explicit MultiLineDelegate(QObject *parent = nullptr)
        : QStyledItemDelegate(parent) {}

    // 创建多行编辑器
    QWidget *createEditor(QWidget *parent, const QStyleOptionViewItem &option,
                          const QModelIndex &index) const override;

    // 将模型数据设置到编辑器
    void setEditorData(QWidget *editor, const QModelIndex &index) const override;

    // 将编辑器数据写回模型
    void setModelData(QWidget *editor, QAbstractItemModel *model,
                      const QModelIndex &index) const override;

    // 调整编辑器几何位置
    void updateEditorGeometry(QWidget *editor, const QStyleOptionViewItem &option,
                              const QModelIndex &index) const override;
};

#endif // MULTILINEDELEGATE_H
