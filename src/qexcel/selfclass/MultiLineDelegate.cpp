#include "MultiLineDelegate.h"
#include <QPlainTextEdit>
#include <QScrollBar>

// 创建多行编辑器（QPlainTextEdit）
QWidget *MultiLineDelegate::createEditor(QWidget *parent, const QStyleOptionViewItem &option,
                                         const QModelIndex &index) const
{
    QPlainTextEdit *editor = new QPlainTextEdit(parent);

    // 设置按编辑器宽度自动换行
    editor->setLineWrapMode(QPlainTextEdit::WidgetWidth);

    // 关键：禁用滚动条
    editor->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    editor->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);

    // 可选：隐藏滚动条（双重保险）
//    editor->verticalScrollBar()->hide();
//    editor->horizontalScrollBar()->hide();

    return editor;
}

// 从模型读取数据到编辑器
void MultiLineDelegate::setEditorData(QWidget *editor, const QModelIndex &index) const
{
    // 获取模型中的原始文本（包含换行符）
    QString text = index.model()->data(index, Qt::EditRole).toString();
    QPlainTextEdit *plainTextEdit = static_cast<QPlainTextEdit*>(editor);
    plainTextEdit->setPlainText(text);
}

// 将编辑器数据写回模型
void MultiLineDelegate::setModelData(QWidget *editor, QAbstractItemModel *model,
                                     const QModelIndex &index) const
{
    QPlainTextEdit *plainTextEdit = static_cast<QPlainTextEdit*>(editor);
    QString text = plainTextEdit->toPlainText();
    model->setData(index, text, Qt::EditRole);
}

// 让编辑器大小匹配单元格
void MultiLineDelegate::updateEditorGeometry(QWidget *editor, const QStyleOptionViewItem &option,
                                             const QModelIndex &index) const
{
    editor->setGeometry(option.rect);
}
