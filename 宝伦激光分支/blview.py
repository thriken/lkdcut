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
    QHeaderView, QFileDialog, QMessageBox, QGroupBox, QFrame, QScrollArea
)
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtNetwork import QNetworkAccessManager, QNetworkRequest, QNetworkReply
import sqlite3


class GFileViewer(QMainWindow):
    """宝伦G文件查看器"""

    def __init__(self):
        super().__init__()
        self.current_file = None
        self.current_folder = None  # 当前文件夹
        self.p3000_data = {}  # P300x数据
        self.p4000_data = []  # P40xx数据
        self.all_pieces_data = []  # 所有文件的小片数据列表
        self.db_conn = None  # 数据库连接
        self.network_manager = QNetworkAccessManager()  # 网络管理器,用于下载图片

        self.init_db_connection()
        self.setup_ui()

    def init_db_connection(self):
        """初始化数据库连接"""
        db_file = os.path.join(os.path.dirname(__file__), '3c.db')
        if os.path.exists(db_file):
            try:
                self.db_conn = sqlite3.connect(db_file)
            except Exception as e:
                pass

    def closeEvent(self, event):
        """窗口关闭时关闭数据库连接"""
        if self.db_conn:
            self.db_conn.close()
        event.accept()

    def load_image_from_url(self, image_url):
        """从URL加载图片"""
        if not image_url or image_url == '-':
            self.image_preview_label.setText("暂无图片")
            self.image_preview_label.setPixmap(QPixmap())
            return

        try:
            url = QUrl(image_url)
            request = QNetworkRequest(url)
            reply = self.network_manager.get(request)

            def on_image_downloaded():
                try:
                    pixmap = QPixmap()
                    pixmap.loadFromData(reply.readAll())
                    if not pixmap.isNull():
                        # 缩放图片以适应显示区域
                        scaled_pixmap = pixmap.scaled(
                            self.image_preview_label.size(),
                            Qt.KeepAspectRatio,
                            Qt.SmoothTransformation
                        )
                        self.image_preview_label.setPixmap(scaled_pixmap)
                    else:
                        self.image_preview_label.setText("图片加载失败")
                except Exception as e:
                    self.image_preview_label.setText(f"图片加载错误: {str(e)}")
                reply.deleteLater()

            reply.finished.connect(on_image_downloaded)
            self.image_preview_label.setText("正在加载图片...")

        except Exception as e:
            self.image_preview_label.setText(f"URL错误: {str(e)}")

    def setup_ui(self):
        """设置界面"""
        self.setWindowTitle("宝伦激光G文件查看器")
        # 窗口大小调整为适配1360*768
        self.setGeometry(100, 100, 1300, 720)

        # 主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # 主布局
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # 顶部工具栏
        toolbar = QHBoxLayout()
        self.create_toolbar(toolbar)
        main_layout.addLayout(toolbar)

        # 创建分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)

        # 小片信息表格
        piece_group = self.create_piece_info_group()
        main_layout.addWidget(piece_group)

    def create_toolbar(self, layout):
        """创建工具栏"""
        open_btn = QPushButton("打开文件")
        open_btn.setFixedWidth(100)
        open_btn.setFont(QFont("Arial", 10))
        open_btn.clicked.connect(self.open_file)
        layout.addWidget(open_btn)

        open_folder_btn = QPushButton("读取文件夹")
        open_folder_btn.setFixedWidth(100)
        open_folder_btn.setFont(QFont("Arial", 10))
        open_folder_btn.clicked.connect(self.open_folder)
        layout.addWidget(open_folder_btn)

        self.file_label = QLabel("未选择文件")
        self.file_label.setStyleSheet("color: gray; font-size: 11px;")
        layout.addWidget(self.file_label)

        layout.addStretch()

    def create_sheet_info_group(self):
        """创建原片信息组"""
        group = QGroupBox("原片参数 (P300x)")
        group.setFont(QFont("Arial", 10, QFont.Bold))

        layout = QVBoxLayout()
        layout.setSpacing(8)

        # 创建字段变量
        self.sheet_vars = {
            'P3000': QLabel(),
            'P3001': QLabel(),
            'P3007': QLabel(),
            'P3011': QLabel(),
        }
        self.piece_seq_label = QLabel()  # 片序号标签

        # 字段显示配置
        field_config = [
            ('P3000', '原片宽度'),
            ('P3001', '原片高度'),
            ('P3007', '原片名称'),
            ('P3011', '原片厚度'),
            ('seq', '小片序号', self.piece_seq_label),  # 特殊处理片序号
        ]

        # 使用统一的布局显示原片信息
        grid_layout = QVBoxLayout()
        grid_layout.setSpacing(8)

        for item in field_config:
            row_layout = QHBoxLayout()
            row_layout.setSpacing(10)

            if len(item) == 3:
                # 特殊项: 片序号
                key, label_text, widget = item
                label = QLabel(f"{label_text}:")
                label.setFont(QFont("Arial", 9))
                label.setMinimumWidth(80)
                row_layout.addWidget(label)

                widget.setFixedWidth(100)
                widget.setFont(QFont("Arial", 9))
                row_layout.addWidget(widget)
            else:
                # 普通项: 原片参数
                key, label_text = item
                label = QLabel(f"{label_text}:")
                label.setFont(QFont("Arial", 9))
                label.setMinimumWidth(80)
                row_layout.addWidget(label)

                edit = self.sheet_vars[key]
                edit.setFixedWidth(100)
                edit.setFont(QFont("Arial", 9))
                row_layout.addWidget(edit)

            row_layout.addStretch()
            grid_layout.addLayout(row_layout)

        layout.addLayout(grid_layout)
        layout.addStretch()
        group.setLayout(layout)
        return group

    def create_piece_info_group(self):
        """创建小片信息组"""
        group = QGroupBox("小片信息")
        group.setFont(QFont("Arial", 10, QFont.Bold))

        layout = QVBoxLayout()
        layout.setSpacing(5)

        # 统计标签
        self.stats_label = QLabel("共 0 个小片")
        self.stats_label.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(self.stats_label)

        # 创建水平布局: 左侧表格, 右侧3C信息
        main_split = QHBoxLayout()
        main_split.setSpacing(10)

        # 左侧: 表格 - 自动扩展
        table_container = QWidget()
        table_layout = QVBoxLayout(table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)

        self.table = QTableWidget()
        self.setup_table()

        # 设置行选择事件
        self.table.itemSelectionChanged.connect(self.on_selection_changed)

        table_layout.addWidget(self.table)
        main_split.addWidget(table_container)  # 不设置stretch,会自动扩展

        # 右侧: 信息区 - 固定宽度320px
        c3_container = QWidget()
        c3_container.setFixedWidth(320)
        c3_layout = QVBoxLayout(c3_container)
        c3_layout.setContentsMargins(0, 0, 0, 0)

        # 创建原片信息显示区域(放于上方)
        sheet_info_group = self.create_sheet_info_group()
        c3_layout.addWidget(sheet_info_group)

        # 创建3C信息显示区域(放于下方)
        c3_info_group = QGroupBox("3C料号信息")
        c3_info_group.setFont(QFont("Arial", 10, QFont.Bold))

        c3_info_layout = QVBoxLayout()
        c3_info_layout.setSpacing(8)

        # 项目名
        project_layout = QHBoxLayout()
        project_layout.addWidget(QLabel("项目名:"))
        self.project_label = QLabel("-")
        self.project_label.setFont(QFont("Arial", 11))
        project_layout.addWidget(self.project_label)
        project_layout.addStretch()
        c3_info_layout.addLayout(project_layout)

        # 标类型
        label_type_layout = QHBoxLayout()
        label_type_layout.addWidget(QLabel("标类型:"))
        self.label_type_label = QLabel("-")
        self.label_type_label.setFont(QFont("Arial", 11))
        label_type_layout.addWidget(self.label_type_label)
        label_type_layout.addStretch()
        c3_info_layout.addLayout(label_type_layout)

        # 备注
        remark_layout = QHBoxLayout()
        remark_layout.addWidget(QLabel("备注:"))
        self.remark_label = QLabel("-")
        self.remark_label.setFont(QFont("Arial", 11))
        remark_layout.addWidget(self.remark_label)
        remark_layout.addStretch()
        c3_info_layout.addLayout(remark_layout)

        # 图片URL - 改为单行显示
        image_url_layout = QHBoxLayout()
        image_url_layout.addWidget(QLabel("图片链接:"))
        self.image_url_label = QLabel("-")
        self.image_url_label.setFont(QFont("Arial", 9))
        self.image_url_label.setWordWrap(False)  # 单行显示
        self.image_url_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.image_url_label.setStyleSheet("color: blue; text-decoration: underline;")
        image_url_layout.addWidget(self.image_url_label)
        image_url_layout.addStretch()
        c3_info_layout.addLayout(image_url_layout)

        # 图片显示区域
        image_group = QGroupBox("图片预览")
        image_group.setFont(QFont("Arial", 9, QFont.Bold))
        image_layout = QVBoxLayout()
        self.image_preview_label = QLabel("暂无图片")
        self.image_preview_label.setAlignment(Qt.AlignCenter)
        self.image_preview_label.setMinimumSize(250, 250)
        self.image_preview_label.setStyleSheet("border: 1px solid #ccc; background-color: #f5f5f5;")
        self.image_preview_label.setScaledContents(False)
        image_layout.addWidget(self.image_preview_label)
        image_group.setLayout(image_layout)
        c3_info_layout.addWidget(image_group)

        c3_info_layout.addStretch()  # 添加弹性空间
        c3_info_group.setLayout(c3_info_layout)
        c3_layout.addWidget(c3_info_group)

        main_split.addWidget(c3_container)  # 添加固定宽度的右侧区域

        layout.addLayout(main_split)

        group.setLayout(layout)
        return group

    def setup_table(self):
        """设置表格"""
        # 定义列 - 添加文件名列和BL切割尺寸列
        columns = [
            '文件名',
            '客户名\n(宝伦7)',
            '货架号\n(宝伦9)',
            '订单尺寸\n(宝伦10)',
            '订单标记\n(宝伦11)',
            '条码号\n(宝伦15)',
            'BL尺寸X\n(宝伦1)',
            'BL尺寸Y\n(宝伦2)',
            '3C基准边\n(宝伦18)',
            '3C料号\n(宝伦17)',
            '3C角位号\n(宝伦19)',
        ]

        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)

        # 设置列宽 - 适配1360*768
        column_widths = [150, 90, 90, 100, 100, 110, 70, 70, 70, 80, 70]
        for i, width in enumerate(column_widths):
            self.table.setColumnWidth(i, width)

        # 设置表格属性
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)  # 单选模式
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # 设置为只读
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(True)  # 显示垂直表头(自带序号)

        # 设置字体 - 减小到10px
        font = QFont("Arial", 10)
        self.table.setFont(font)
        self.table.horizontalHeader().setFont(QFont("Arial", 10, QFont.Bold))
        self.table.verticalHeader().setFont(QFont("Arial", 9))

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

    def open_folder(self):
        """读取文件夹中的所有G文件"""
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择包含G文件的文件夹",
            ""
        )

        if folder:
            self.parse_folder(folder)

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
            self.current_folder = None
            self.file_label.setText(os.path.basename(filepath))
            self.file_label.setStyleSheet("color: black;")

            # 清空之前的数据
            self.p3000_data = {}
            self.p4000_data = []
            self.all_pieces_data = []

            # 解析文件名获取原片序号
            # 文件名格式: 订单单号_原片序号_片数量_原片长_原片高.g
            # 例如: YH260307058_003_1_3660_3000.g
            sheet_seq = self.parse_sheet_sequence(os.path.basename(filepath))

            # 解析文件
            piece_seq = 0  # 片序号计数器
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
                    piece_data['文件名'] = os.path.basename(filepath)
                    piece_data['原片序号'] = sheet_seq  # 保存原片序号
                    piece_data['片序号'] = piece_seq + 1  # 保存片序号(从1开始)
                    piece_seq += 1
                    self.p4000_data.append(piece_data)

            # 更新界面
            self.clear_sheet_info()  # 清空原片信息,等待选择后加载
            self.update_piece_info()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"解析文件失败: {str(e)}")

    def parse_folder(self, folder):
        """解析文件夹中的所有G文件"""
        try:
            self.current_folder = folder
            self.current_file = None
            self.file_label.setText(f"文件夹: {os.path.basename(folder)} ({len(os.listdir(folder))} 个文件)")
            self.file_label.setStyleSheet("color: black;")

            # 清空之前的数据
            self.p3000_data = {}
            self.p4000_data = []
            self.all_pieces_data = []

            # 遍历文件夹中的所有文件
            files = sorted([f for f in os.listdir(folder) if f.lower().endswith('.g')])

            for filename in files:
                filepath = os.path.join(folder, filename)
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
                        continue

                    # 解析文件名获取原片序号
                    # 文件名格式: 订单单号_原片序号_片数量_原片长_原片高.g
                    # 例如: YH260307058_003_1_3660_3000.g
                    sheet_seq = self.parse_sheet_sequence(filename)

                    # 解析文件
                    file_p4000_data = []
                    file_sheet_data = {}  # 当前文件的原片信息
                    piece_seq = 0  # 每个文件独立的片序号
                    for line in lines:
                        line = line.strip()
                        if not line:
                            continue

                        # 解析P300x参数(当前文件的原片参数)
                        match_p300 = re.match(r'^N\d{2}\s+P(30\d{2})\s*=\s*(.*)', line)
                        if match_p300:
                            param_name = f'P{match_p300.group(1)}'
                            param_value = match_p300.group(2).strip()
                            file_sheet_data[param_name] = param_value

                        # 解析P40xx小片数据
                        match_p400 = re.match(r'^N\d{2}\s+P(4\d{3})\s*=\s*(.+)', line)
                        if match_p400:
                            param_num = int(match_p400.group(1))
                            data_string = match_p400.group(2).strip()
                            piece_data = self.parse_piece_data(data_string, param_num)
                            piece_data['文件名'] = filename
                            piece_data['原片序号'] = sheet_seq  # 保存原片序号
                            piece_data['片序号'] = piece_seq + 1  # 保存片序号
                            piece_data['原片信息'] = file_sheet_data.copy()  # 保存该文件的原片信息
                            piece_seq += 1
                            file_p4000_data.append(piece_data)

                    self.all_pieces_data.extend(file_p4000_data)

                except Exception as e:
                    continue  # 跳过有问题的文件

            # 更新界面
            self.clear_sheet_info()  # 清空原片信息,等待选择后加载
            self.update_piece_info_from_folder()

        except Exception as e:
            QMessageBox.critical(self, "错误", f"解析文件夹失败: {str(e)}")

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
        self.piece_seq_label.setText('-')

    def clear_sheet_info(self):
        """清空原片信息显示"""
        self.sheet_vars['P3000'].setText('')
        self.sheet_vars['P3001'].setText('')
        self.sheet_vars['P3007'].setText('')
        self.sheet_vars['P3011'].setText('')
        self.piece_seq_label.setText('-')

    def parse_sheet_sequence(self, filename):
        """从文件名解析原片序号

        文件名格式: 订单单号_原片序号_片数量_原片长_原片高.g
        例如: YH260307058_003_1_3660_3000.g

        Args:
            filename: G文件名

        Returns:
            原片序号,如果解析失败返回'-'
        """
        # 移除扩展名
        basename = os.path.splitext(filename)[0]
        parts = basename.split('_')

        # 至少需要5个部分: 订单单号_原片序号_片数量_原片长_原片高
        if len(parts) >= 5:
            # 第二部分是原片序号
            sheet_seq = parts[1]
            # 尝试转换为整数以去除前导零
            try:
                sheet_seq_int = int(sheet_seq)
                return str(sheet_seq_int)
            except ValueError:
                return sheet_seq

        return '-'

    def update_piece_info(self):
        """更新小片信息表格"""
        # 清空表格
        self.table.setRowCount(0)

        # 设置行数
        row_count = len(self.p4000_data)
        self.table.setRowCount(row_count)

        # 添加数据
        for row_idx, piece in enumerate(self.p4000_data):
            # 显示的字段: 文件名, 7,9,10,11,15,1,2,18,17,19
            self.table.setItem(row_idx, 0, QTableWidgetItem(piece.get('文件名', '')))   # 文件名
            self.table.setItem(row_idx, 1, QTableWidgetItem(piece['宝伦7']))   # 客户名
            self.table.setItem(row_idx, 2, QTableWidgetItem(piece['宝伦9']))   # 货架号
            self.table.setItem(row_idx, 3, QTableWidgetItem(piece['宝伦10']))  # 订单号
            self.table.setItem(row_idx, 4, QTableWidgetItem(piece['宝伦11']))  # 订单尺寸
            self.table.setItem(row_idx, 5, QTableWidgetItem(piece['宝伦15']))  # 条码号
            self.table.setItem(row_idx, 6, QTableWidgetItem(piece['宝伦1']))   # BL尺寸X
            self.table.setItem(row_idx, 7, QTableWidgetItem(piece['宝伦2']))   # BL尺寸Y
            self.table.setItem(row_idx, 8, QTableWidgetItem(piece['宝伦18']))  # 基准边1
            self.table.setItem(row_idx, 9, QTableWidgetItem(piece['宝伦17'])) # 标签1模板
            self.table.setItem(row_idx, 10, QTableWidgetItem(piece['宝伦19'])) # 角位号1

            # 设置单元格对齐
            for col in range(11):
                item = self.table.item(row_idx, col)
                if item:
                    item.setTextAlignment(Qt.AlignCenter)

        # 更新统计信息
        self.stats_label.setText(f"共 {row_count} 个小片")

    def update_piece_info_from_folder(self):
        """从文件夹数据更新小片信息表格"""
        # 清空表格
        self.table.setRowCount(0)

        # 设置行数
        row_count = len(self.all_pieces_data)
        self.table.setRowCount(row_count)

        # 添加数据
        for row_idx, piece in enumerate(self.all_pieces_data):
            # 显示的字段: 文件名, 7,9,10,11,15,1,2,18,17,19
            self.table.setItem(row_idx, 0, QTableWidgetItem(piece.get('文件名', '')))   # 文件名
            self.table.setItem(row_idx, 1, QTableWidgetItem(piece['宝伦7']))   # 客户名
            self.table.setItem(row_idx, 2, QTableWidgetItem(piece['宝伦9']))   # 货架号
            self.table.setItem(row_idx, 3, QTableWidgetItem(piece['宝伦10']))  # 订单号
            self.table.setItem(row_idx, 4, QTableWidgetItem(piece['宝伦11']))  # 订单尺寸
            self.table.setItem(row_idx, 5, QTableWidgetItem(piece['宝伦15']))  # 条码号
            self.table.setItem(row_idx, 6, QTableWidgetItem(piece['宝伦1']))   # BL尺寸X
            self.table.setItem(row_idx, 7, QTableWidgetItem(piece['宝伦2']))   # BL尺寸Y
            self.table.setItem(row_idx, 8, QTableWidgetItem(piece['宝伦18']))  # 基准边1
            self.table.setItem(row_idx, 9, QTableWidgetItem(piece['宝伦17'])) # 标签1模板
            self.table.setItem(row_idx, 10, QTableWidgetItem(piece['宝伦19'])) # 角位号1

            # 设置单元格对齐
            for col in range(11):
                item = self.table.item(row_idx, col)
                if item:
                    item.setTextAlignment(Qt.AlignCenter)

        # 更新统计信息
        self.stats_label.setText(f"共 {row_count} 个小片")



    def on_selection_changed(self):
        """行选择改变事件"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            self.clear_c3_info()
            return

        # 获取选中行的索引(0-based,即内部索引)
        row = selected_items[0].row()

        # 加载原片信息和3C信息
        self.load_selected_piece_info(row)
        self.update_c3_info(row)

    def load_selected_piece_info(self, row):
        """加载选中行的原片信息和片序号"""
        # 获取对应的数据源
        data_source = self.all_pieces_data if self.current_folder else self.p4000_data

        if row < len(data_source):
            piece = data_source[row]

            # 更新片序号: 原片序号-小片序号
            sheet_seq = piece.get('原片序号', '-')
            piece_seq = piece.get('片序号', '-')
            self.piece_seq_label.setText(f"{sheet_seq}-{piece_seq}")

            # 显示原片信息
            if self.current_folder:
                # 文件夹模式: 从对应的小片数据中获取原片信息
                sheet_info = piece.get('原片信息', {})
                self.sheet_vars['P3000'].setText(sheet_info.get('P3000', ''))
                self.sheet_vars['P3001'].setText(sheet_info.get('P3001', ''))
                self.sheet_vars['P3007'].setText(sheet_info.get('P3007', ''))
                self.sheet_vars['P3011'].setText(sheet_info.get('P3011', ''))
            else:
                # 单个文件模式,显示原片信息
                if self.p3000_data:
                    self.sheet_vars['P3000'].setText(self.p3000_data.get('P3000', ''))
                    self.sheet_vars['P3001'].setText(self.p3000_data.get('P3001', ''))
                    self.sheet_vars['P3007'].setText(self.p3000_data.get('P3007', ''))
                    self.sheet_vars['P3011'].setText(self.p3000_data.get('P3011', ''))

    def query_c3_info_from_db(self, material_num):
        """从数据库查询3C料号信息"""
        if not self.db_conn:
            return None

        cursor = self.db_conn.cursor()

        # 查询1: 在jingling_standard列查找(正标)
        cursor.execute("SELECT * FROM materials WHERE jingling_standard = ?", (material_num,))
        row = cursor.fetchone()
        if row:
            # row[2]=project_name, row[8]=image_path
            return {
                '项目名': row[2],  # project_name列(索引2)
                '标类型': '正标',
                '备注': row[7] if len(row) > 7 else '',  # remark(索引7)
                '图片路径': row[8] if len(row) > 8 else '',  # image_path(索引8)
            }

        # 查询2: 在reverse_standard列查找(反标)
        cursor.execute("SELECT * FROM materials WHERE reverse_standard = ?", (material_num,))
        row = cursor.fetchone()
        if row:
            # row[2]=project_name, row[8]=image_path
            return {
                '项目名': row[2],  # project_name列(索引2)
                '标类型': '反标',
                '备注': row[7] if len(row) > 7 else '',  # remark(索引7)
                '图片路径': row[8] if len(row) > 8 else '',  # image_path(索引8)
            }

        # 查询3: 在lowe_special列查找(LOWE专用标)
        cursor.execute("SELECT * FROM materials WHERE lowe_special = ?", (material_num,))
        row = cursor.fetchone()
        if row:
            # row[2]=project_name, row[9]=lowe_image_path
            return {
                '项目名': row[2],  # project_name列(索引2)
                '标类型': 'LOWE专用标',
                '备注': row[7] if len(row) > 7 else '',  # remark(索引7)
                '图片路径': row[9] if len(row) > 9 else '',  # lowe_image_path(索引9)
            }

        return None

    def update_c3_info(self, row):
        """更新3C信息显示"""
        # 获取宝伦17列(索引9)的料号
        item = self.table.item(row, 9)
        if not item:
            self.clear_c3_info()
            return

        material_num = item.text().strip()

        if not material_num:
            self.clear_c3_info()
            return

        # 从数据库查询料号信息
        c3_info = self.query_c3_info_from_db(material_num)

        if c3_info:
            # 更新UI
            self.project_label.setText(c3_info['项目名'])
            self.label_type_label.setText(c3_info['标类型'])
            self.remark_label.setText(c3_info['备注'] if c3_info['备注'] else '-')

            # 构建图片URL并显示
            image_path = c3_info.get('图片路径', '')
            if image_path:
                image_url = f"https://win7e.com/{image_path}"
                self.image_url_label.setText(image_url)
                # 加载图片预览
                self.load_image_from_url(image_url)
            else:
                self.image_url_label.setText('-')
                self.image_preview_label.setText("暂无图片")
                self.image_preview_label.setPixmap(QPixmap())
        else:
            self.clear_c3_info()

    def clear_c3_info(self):
        """清空3C信息显示"""
        self.project_label.setText('-')
        self.label_type_label.setText('-')
        self.remark_label.setText('-')
        self.image_url_label.setText('-')
        self.image_preview_label.setText("暂无图片")
        self.image_preview_label.setPixmap(QPixmap())


def main():
    """主函数"""
    app = QApplication(sys.argv)
    viewer = GFileViewer()
    viewer.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
