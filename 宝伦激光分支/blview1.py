#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
宝伦激光G文件查看程序 - Qt版本
用于检查G文件中的数据信息
"""

import os
import re
import sys
from typing import Dict, List

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QMessageBox, QGroupBox, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class GFileViewer(QMainWindow):
    """宝伦G文件查看器"""

    def __init__(self):
        super().__init__()
        self.current_file = None
        self.p3000_data = {}  # P300x数据
        self.p4000_data = []  # P40xx数据

        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        self.setWindowTitle("宝伦激光G文件查看器")
        self.setGeometry(100, 100, 1400, 800)

        # 主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # 主布局
        main_layout = QVBoxLayout(main_widget)

        # 顶部工具栏
        toolbar = QHBoxLayout()
        self.create_toolbar(toolbar)
        main_layout.addLayout(toolbar)

        # 创建分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)

        # 上部：原片信息
        sheet_group = self.create_sheet_info_group()
        main_layout.addWidget(sheet_group)

        # 下部：小片信息表格
        piece_group = self.create_piece_info_group()
        main_layout.addWidget(piece_group)

    def create_toolbar(self, layout):
        """创建工具栏"""
        open_btn = QPushButton("打开文件")
        open_btn.setFixedWidth(120)
        open_btn.clicked.connect(self.open_file)
        layout.addWidget(open_btn)

        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray;")
        layout.addWidget(self.file_label)

        layout.addStretch()

    def create_sheet_info_group(self):
        """创建原片信息组"""
        group = QGroupBox("原片参数 (P300x)")
        group.setFont(QFont("Arial", 12, QFont.Bold))

        layout = QHBoxLayout()

        # 创建字段变量
        self.sheet_vars = {
            'P3000': QLineEdit(),
            'P3001': QLineEdit(),
            'P3007': QLineEdit(),
            'P3011': QLineEdit(),
        }

        labels = {
            'P3000': '原片宽度 (P3000):',
            'P3001': '原片高度 (P3001):',
            'P3007': '原片名称 (P3007):',
            'P3011': '原片厚度 (P3011):',
        }

        # 使用水平布局显示原片信息
        for key, label_text in labels.items():
            label = QLabel(label_text)
            label.setFont(QFont("Arial", 11))
            label.setMinimumWidth(150)
            layout.addWidget(label)

            edit = self.sheet_vars[key]
            edit.setFixedWidth(120)
            edit.setReadOnly(True)
            edit.setFont(QFont("Arial", 11))
            layout.addWidget(edit)

            layout.addSpacing(20)

        layout.addStretch()
        group.setLayout(layout)
        return group

    def create_piece_info_group(self):
        """创建小片信息组"""
        group = QGroupBox("小片信息 (P40xx)")
        group.setFont(QFont("Arial", 12, QFont.Bold))

        layout = QVBoxLayout()

        # 统计标签
        self.stats_label = QLabel("共 0 个小片")
        self.stats_label.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(self.stats_label)

        # 创建表格
        self.table = QTableWidget()
        self.setup_table()

        layout.addWidget(self.table)
        group.setLayout(layout)
        return group

    def setup_table(self):
        """设置表格"""
        # 定义列 - 使用BL1-BL28显示所有参数
        columns = ['序号'] + [f'BL{i}' for i in range(1, 29)]

        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)

        # 设置列宽 - 前3列较窄，其他列适当宽度
        column_widths = [50] + [80] * 28
        for i, width in enumerate(column_widths):
            self.table.setColumnWidth(i, width)

        # 设置表格属性
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)  # 隐藏行号列

        # 设置字体
        font = QFont("Arial", 8)
        self.table.setFont(font)
        self.table.horizontalHeader().setFont(QFont("Arial", 8, QFont.Bold))

    def open_file(self):
        """打开G文件"""
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "选择宝伦激光G文件",
            "",
            "G文件 (*.g *.G);;所有文件 (*.*)"
        )

        if filepath:
            self.parse_file(filepath)

    def parse_file(self, filepath):
        """解析G文件"""
        try:
            # 尝试不同的编码
            lines = None
            for encoding in ['gb2312', 'utf-8', 'gbk']:
                try:
                    with open(filepath, 'r', encoding=encoding, errors='ignore') as f:
                        lines = f.readlines()
                    break
                except:
                    continue

            if lines is None:
                QMessageBox.critical(self, "错误", "无法读取文件！")
                return

            self.current_file = filepath
            self.file_label.setText(os.path.basename(filepath))
            self.file_label.setStyleSheet("color: black;")

            # 清空之前的数据
            self.p3000_data = {}
            self.p4000_data = []

            # 解析文件
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # 解析P300x参数
                match_p300 = re.match(r'^N\d{2}\s+P(30\d{2})\s*=\s*(.*)', line)
                if match_p300:
                    param_name = f'P{match_p300.group(1)}'
                    param_value = match_p300.group(2).strip()
                    self.p3000_data[param_name] = param_value
                    continue

                # 解析P40xx小片数据
                match_p400 = re.match(r'^N\d{2}\s+P(4\d{3})\s*=\s*(.+)', line)
                if match_p400:
                    param_num = int(match_p400.group(1))
                    data_string = match_p400.group(2).strip()
                    piece_data = self.parse_piece_data(data_string, param_num)
                    self.p4000_data.append(piece_data)

            # 更新界面
            self.update_sheet_info()
            self.update_piece_info()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"解析文件失败: {str(e)}")

    def parse_piece_data(self, data_string: str, param_num: int) -> Dict:
        """解析P40xx小片数据"""
        parts = data_string.split('_')

        # 宝伦字段有28个
        piece_info = {
            'param_num': param_num,
            '宝伦1': parts[0] if len(parts) > 0 else '',   # X方向尺寸
            '宝伦2': parts[1] if len(parts) > 1 else '',   # Y方向尺寸
            '宝伦3': parts[2] if len(parts) > 2 else '',   # 左下角X坐标
            '宝伦4': parts[3] if len(parts) > 3 else '',   # 左下角Y坐标
            '宝伦5': parts[4] if len(parts) > 4 else '',   # 显示的长(X)
            '宝伦6': parts[5] if len(parts) > 5 else '',   # 显示的宽(Y)
            '宝伦7': parts[6] if len(parts) > 6 else '',   # 客户名
            '宝伦8': parts[7] if len(parts) > 7 else '',   # 定位号
            '宝伦9': parts[8] if len(parts) > 8 else '',   # 货架号分组号
            '宝伦10': parts[9] if len(parts) > 9 else '',  # 订单号
            '宝伦11': parts[10] if len(parts) > 10 else '', # 订单尺寸加片标记
            '宝伦12': parts[11] if len(parts) > 11 else '', # 空
            '宝伦13': parts[12] if len(parts) > 12 else '', # 空
            '宝伦14': parts[13] if len(parts) > 13 else '', # 空
            '宝伦15': parts[14] if len(parts) > 14 else '', # 条码号
            '宝伦16': parts[15] if len(parts) > 15 else '', # 小片数量
            '宝伦17': parts[16] if len(parts) > 16 else '', # 标签1激光模板序号
            '宝伦18': parts[17] if len(parts) > 17 else '', # 标签1基准边
            '宝伦19': parts[18] if len(parts) > 18 else '', # 标签1角位号
            '宝伦20': parts[19] if len(parts) > 19 else '', # 标签1离边角位置
            '宝伦21': parts[20] if len(parts) > 20 else '', # 标签2激光模板序号
            '宝伦22': parts[21] if len(parts) > 21 else '', # 标签2基准边
            '宝伦23': parts[22] if len(parts) > 22 else '', # 标签2角位号
            '宝伦24': parts[23] if len(parts) > 23 else '', # 标签2离边角位置
            '宝伦25': parts[24] if len(parts) > 24 else '', # 标签3模板序号
            '宝伦26': parts[25] if len(parts) > 25 else '', # 标签3基准边
            '宝伦27': parts[26] if len(parts) > 26 else '', # 标签3角位号
            '宝伦28': parts[27] if len(parts) > 27 else '', # 标签3离边角位置
        }

        return piece_info

    def update_sheet_info(self):
        """更新原片信息显示"""
        self.sheet_vars['P3000'].setText(self.p3000_data.get('P3000', ''))
        self.sheet_vars['P3001'].setText(self.p3000_data.get('P3001', ''))
        self.sheet_vars['P3007'].setText(self.p3000_data.get('P3007', ''))
        self.sheet_vars['P3011'].setText(self.p3000_data.get('P3011', ''))

    def update_piece_info(self):
        """更新小片信息表格"""
        # 清空表格
        self.table.setRowCount(0)

        # 设置行数
        row_count = len(self.p4000_data)
        self.table.setRowCount(row_count)

        # 添加数据 - 显示所有28个参数
        for row_idx, piece in enumerate(self.p4000_data):
            # 序号
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(row_idx + 1)))

            # 显示所有28个BL参数
            for i in range(1, 29):
                bl_key = f'宝伦{i}'
                self.table.setItem(row_idx, i, QTableWidgetItem(piece[bl_key]))

            # 设置单元格对齐
            for col in range(29):
                item = self.table.item(row_idx, col)
                if item:
                    item.setTextAlignment(Qt.AlignCenter)

        # 更新统计信息
        self.stats_label.setText(f"共 {row_count} 个小片")


def main():
    """主函数"""
    app = QApplication(sys.argv)
    viewer = GFileViewer()
    viewer.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
