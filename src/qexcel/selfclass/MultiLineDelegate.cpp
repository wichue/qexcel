#include "MultiLineDelegate.h"
#include <QtSql>

// 创建多行编辑器（QPlainTextEdit）
QWidget *MultiLineDelegate::createEditor(QWidget *parent, const QStyleOptionViewItem &option,
                                         const QModelIndex &index) const
{
    QPlainTextEdit *editor = new QPlainTextEdit(parent);
    // 设置按编辑器宽度自动换行，避免横向滚动
    editor->setLineWrapMode(QPlainTextEdit::WidgetWidth);
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
