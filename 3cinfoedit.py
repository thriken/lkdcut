"""
3C信息编辑工具 - 管理cccindex表
"""
import sys
import sqlite3
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QLineEdit, QLabel,
    QMessageBox, QDialog, QFormLayout, QDialogButtonBox, QHeaderView,
    QGroupBox, QComboBox
)
from PyQt5.QtCore import Qt



class CCCIndexEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db_path = r"e:\Project\lkdcut\3c.db"
        self.init_ui()
        self.load_data()

    def init_ui(self):
        self.setWindowTitle("3C信息编辑工具 - cccindex表管理")
        self.setGeometry(100, 100, 1200, 800)

        # 中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        # 顶部工具栏
        toolbar = QHBoxLayout()
        
        # 搜索框
        toolbar.addWidget(QLabel("搜索:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入laserid或项目名称搜索...")
        self.search_input.textChanged.connect(self.on_search)
        toolbar.addWidget(self.search_input)

        toolbar.addStretch()

        # 按钮
        self.btn_add = QPushButton("新增")
        self.btn_add.clicked.connect(self.add_record)
        toolbar.addWidget(self.btn_add)

        self.btn_edit = QPushButton("编辑")
        self.btn_edit.clicked.connect(self.edit_record)
        toolbar.addWidget(self.btn_edit)

        self.btn_delete = QPushButton("删除")
        self.btn_delete.clicked.connect(self.delete_record)
        toolbar.addWidget(self.btn_delete)

        self.btn_refresh = QPushButton("刷新")
        self.btn_refresh.clicked.connect(self.load_data)
        toolbar.addWidget(self.btn_refresh)

        layout.addLayout(toolbar)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "ID", "laserid", "process_num", "customer", "project_name", "remark", "type"
        ])
        # 设置列宽: 前两列固定宽度，后面的列自适应
        self.table.setColumnWidth(0, 50)   # ID
        self.table.setColumnWidth(1, 60)   # laserid
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)  # process_num
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)  # customer
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)  # project_name
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)  # remark
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)  # type
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.doubleClicked.connect(self.edit_record)
        layout.addWidget(self.table)

        # 状态栏
        self.status_label = QLabel("就绪")
        layout.addWidget(self.status_label)

        central_widget.setLayout(layout)

    def get_connection(self):
        return sqlite3.connect(self.db_path)

    def load_data(self, search_text=""):
        """加载数据"""
        conn = self.get_connection()
        cursor = conn.cursor()

        if search_text:
            cursor.execute("""
                SELECT id, laserid, process_num, customer, project_name, remark, type
                FROM cccindex
                WHERE laserid LIKE ? OR project_name LIKE ? OR process_num LIKE ? OR remark LIKE ?
                ORDER BY id
            """, (f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"))
        else:
            cursor.execute("""
                SELECT id, laserid, process_num, customer, project_name, remark, type
                FROM cccindex
                ORDER BY id
            """)

        rows = cursor.fetchall()
        conn.close()

        self.table.setRowCount(len(rows))
        for row_idx, row_data in enumerate(rows):
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem(str(value) if value else "")
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, col_idx, item)

        self.status_label.setText(f"共 {len(rows)} 条记录")

    def on_search(self, text):
        """搜索"""
        self.load_data(text)

    def get_selected_id(self):
        """获取选中的记录ID"""
        selected_rows = self.table.selectedItems()
        if not selected_rows:
            return None
        row = selected_rows[0].row()
        return int(self.table.item(row, 0).text())

    def get_record_by_id(self, record_id):
        """根据ID获取记录"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, laserid, process_num, customer, project_name, remark, type
            FROM cccindex WHERE id = ?
        """, (record_id,))
        row = cursor.fetchone()
        conn.close()
        return row

    def add_record(self):
        """新增记录"""
        dialog = RecordDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO cccindex (laserid, process_num, customer, project_name, remark, type)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (data['laserid'], data['process_num'], data['customer'], 
                  data['project_name'], data['remark'], data['type']))
            conn.commit()
            conn.close()
            self.load_data(self.search_input.text())
            self.status_label.setText("新增成功")

    def edit_record(self):
        """编辑记录"""
        record_id = self.get_selected_id()
        if not record_id:
            QMessageBox.warning(self, "提示", "请先选择要编辑的记录")
            return

        record = self.get_record_by_id(record_id)
        dialog = RecordDialog(self, record)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE cccindex
                SET laserid=?, process_num=?, customer=?, project_name=?, remark=?, type=?
                WHERE id=?
            """, (data['laserid'], data['process_num'], data['customer'],
                  data['project_name'], data['remark'], data['type'], record_id))
            conn.commit()
            conn.close()
            self.load_data(self.search_input.text())
            self.status_label.setText("修改成功")

    def delete_record(self):
        """删除记录"""
        record_id = self.get_selected_id()
        if not record_id:
            QMessageBox.warning(self, "提示", "请先选择要删除的记录")
            return

        reply = QMessageBox.question(self, "确认删除", "确定要删除选中的记录吗？",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            conn = self.get_connection()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM cccindex WHERE id=?", (record_id,))
            conn.commit()
            conn.close()
            self.load_data(self.search_input.text())
            self.status_label.setText("删除成功")


class RecordDialog(QDialog):
    """记录编辑对话框"""
    def __init__(self, parent=None, record=None):
        super().__init__(parent)
        self.record = record
        self.init_ui()

    def init_ui(self):
        if self.record:
            self.setWindowTitle("编辑记录")
        else:
            self.setWindowTitle("新增记录")

        layout = QFormLayout()

        # 字段输入框
        self.laserid_input = QLineEdit()
        self.process_num_input = QLineEdit()
        self.customer_input = QLineEdit()
        self.project_name_input = QLineEdit()
        self.remark_input = QLineEdit()

        self.type_input = QComboBox()
        self.type_input.addItems(["正标", "反标", "LOWE标", ""])

        # 如果是编辑模式，填充数据
        if self.record:
            self.laserid_input.setText(str(self.record[1]) if self.record[1] else "")
            self.process_num_input.setText(str(self.record[2]) if self.record[2] else "")
            self.customer_input.setText(str(self.record[3]) if self.record[3] else "")
            self.project_name_input.setText(str(self.record[4]) if self.record[4] else "")
            self.remark_input.setText(str(self.record[5]) if self.record[5] else "")
            type_val = str(self.record[6]) if self.record[6] else ""
            index = self.type_input.findText(type_val)
            if index >= 0:
                self.type_input.setCurrentIndex(index)

        layout.addRow("laserid:", self.laserid_input)
        layout.addRow("process_num:", self.process_num_input)
        layout.addRow("customer:", self.customer_input)
        layout.addRow("project_name:", self.project_name_input)
        layout.addRow("remark:", self.remark_input)
        layout.addRow("type:", self.type_input)

        # 按钮
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

        self.setLayout(layout)

    def get_data(self):
        """获取输入的数据"""
        return {
            'laserid': self.laserid_input.text().strip(),
            'process_num': self.process_num_input.text().strip(),
            'customer': self.customer_input.text().strip(),
            'project_name': self.project_name_input.text().strip(),
            'remark': self.remark_input.text().strip(),
            'type': self.type_input.currentText()
        }


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei UI", 12))
    window = CCCIndexEditor()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
