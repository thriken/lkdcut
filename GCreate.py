# -*- coding: utf-8 -*-
"""
G文件创建工具 - 根据模板生成G代码文件
"""
import os
import sys
import sqlite3
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog,
    QGroupBox, QFormLayout, QComboBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap


class GCreteWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_material = None
        self.init_ui()
        
        # G代码导出配置
        self.GCODE_EXPORT_DIR = r"D:\Program Files\NewGlass\DrillWork\\"
    
    def init_ui(self):
        self.setWindowTitle("G文件创建工具")
        self.setGeometry(100, 100, 900, 930)
        
        # 设置窗口样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #e0e0e0;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #1890ff;
            }
            QLabel {
                color: #595959;
                font-size: 14px;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background: #fafafa;
                min-width: 120px;
            }
            QLineEdit:focus {
                border-color: #40a9ff;
            }
            QComboBox {
                padding: 8px 12px;
                border: 1px solid #d9d9d9;
                border-radius: 6px;
                background: #ffffff;
                min-width: 100px;
                font-size: 14px;
                color: #333;
            }
            QComboBox:hover {
                border-color: #40a9ff;
            }
            QComboBox:focus {
                border-color: #1890ff;
                background: #fff;
            }
            QComboBox::drop-down {
                border: none;
                width: 30px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #999;
                margin-right: 8px;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #d9d9d9;
                border-radius: 6px;
                background: #ffffff;
                selection-background-color: #e6f7ff;
                selection-color: #1890ff;
                font-size: 14px;
                padding: 4px;
            }
            QPushButton {
                min-width: 100px;
                padding: 8px 15px;
                background-color: #1890ff;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: 500;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
            QPushButton:pressed {
                background-color: #096dd9;
            }
            QTextEdit {
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background: #fafafa;
                font-family: Consolas, Monaco, monospace;
                font-size: 12px;
            }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # ==================== 上层：玻璃参数和图片预览 1:1 ====================
        top_layout = QHBoxLayout()
        top_layout.setSpacing(15)

        # 左侧：玻璃参数区域
        glass_group = QGroupBox("玻璃参数")
        glass_layout = QFormLayout()
        glass_layout.setLabelAlignment(Qt.AlignRight)
        glass_layout.setSpacing(15)
        
        # 目标尺寸
        size_layout = QHBoxLayout()
        self.target_x = QLineEdit()
        self.target_x.setPlaceholderText("X")
        self.target_x.setMaximumWidth(80)
        self.target_y = QLineEdit()
        self.target_y.setPlaceholderText("Y")
        self.target_y.setMaximumWidth(80)
        size_layout.addWidget(QLabel(""))
        size_layout.addWidget(self.target_x)
        size_layout.addWidget(QLabel("×"))
        size_layout.addWidget(self.target_y)
        size_layout.addStretch()
        glass_layout.addRow("目标尺寸:", size_layout)
        
        # 磨边量
        self.edge_combo = QComboBox()
        self.edge_combo.addItems(['0', '1', '2'])
        glass_layout.addRow("磨边量:", self.edge_combo)
        
        # 原片厚度
        self.thickness_combo = QComboBox()
        self.thickness_combo.addItems(['4', '5', '6', '8', '10', '12', '15'])
        self.thickness_combo.setCurrentText('5')
        glass_layout.addRow("原片厚度:", self.thickness_combo)
        
        # 原片类型
        self.glass_type_combo = QComboBox()
        self.glass_type_combo.addItems(['白玻', '玉', 'LOWE'])
        glass_layout.addRow("原片类型:", self.glass_type_combo)
        
        # 料号输入
        self.material_edit = QLineEdit()
        self.material_edit.setPlaceholderText("输入料号...")
        self.material_edit.editingFinished.connect(self.on_material_changed)
        glass_layout.addRow("料号:", self.material_edit)
        
        # 3C码位置
        self.pos_3c = QComboBox()
        self.pos_3c.addItems(['1=左下', '2=左上', '3=右下', '4=右上'])
        glass_layout.addRow("3C码位置:", self.pos_3c)
        
        glass_group.setLayout(glass_layout)
        top_layout.addWidget(glass_group, 1)  # 1:1比例

        # 右侧：图片预览区域
        img_group = QGroupBox("料号图片预览")
        img_layout = QVBoxLayout()
        img_layout.setContentsMargins(10, 10, 10, 10)

        self.img_label = QLabel("请输入料号查看图片")
        self.img_label.setAlignment(Qt.AlignCenter)
        self.img_label.setMinimumSize(280, 280)
        self.img_label.setStyleSheet("""
            border: 2px dashed #d9d9d9;
            background: #fafafa;
            color: #8c8c8c;
            font-size: 14px;
        """)
        img_layout.addWidget(self.img_label)

        img_group.setLayout(img_layout)
        top_layout.addWidget(img_group, 1)  # 1:1比例

        main_layout.addLayout(top_layout)

        # ==================== 下层：G代码预览区域 独占全部宽度 ====================
        preview_group = QGroupBox("G代码预览")
        preview_layout = QVBoxLayout()
        preview_layout.setContentsMargins(5, 5, 5, 5)
        self.preview_edit = QTextEdit()
        self.preview_edit.setReadOnly(True)
        self.preview_edit.setMinimumHeight(200)
        self.preview_edit.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background: #1e1e1e;
                color: #d4d4d4;
                font-family: Consolas, Monaco, monospace;
                font-size: 12px;
                padding: 8px;
            }
        """)
        preview_layout.addWidget(self.preview_edit)
        preview_group.setLayout(preview_layout)
        main_layout.addWidget(preview_group)

        # ==================== 按钮区域 ====================
        btn_layout = QHBoxLayout()

        btn_preview = QPushButton("生成预览")
        btn_preview.clicked.connect(self.preview_gcode)
        btn_layout.addWidget(btn_preview)

        btn_export = QPushButton("导出G文件")
        btn_export.clicked.connect(self.export_gcode)
        btn_layout.addWidget(btn_export)

        btn_layout.addStretch()
        main_layout.addLayout(btn_layout)

    def on_material_changed(self):
        """料号输入变化时加载对应图片"""
        material_id = self.material_edit.text().strip()
        if not material_id:
            self.img_label.setText("请输入料号\n查看图片预览")
            self.img_label.setStyleSheet("""
                border: 2px dashed #d9d9d9;
                background: #fafafa;
                color: #8c8c8c;
                font-size: 14px;
            """)
            self.selected_material = None
            return
        
        # 构建图片路径
        img_path = os.path.join(os.path.dirname(__file__), 'resources', 'img', f'{material_id}.png')
        
        # 尝试加载图片
        if os.path.exists(img_path):
            pixmap = QPixmap(img_path)
            if not pixmap.isNull():
                # 缩放图片以适应显示区域
                scaled_pixmap = pixmap.scaled(260, 260, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.img_label.setPixmap(scaled_pixmap)
                self.img_label.setStyleSheet("border: 2px solid #52c41a; background: #f6ffed;")
                print(f"加载图片成功: {img_path}")
            else:
                self.img_label.setText("图片加载失败")
                self.img_label.setStyleSheet("border: 2px solid #ff4d4f; background: #fff2f0;")
        else:
            self.img_label.setText(f"图片不存在\n{material_id}.png")
            self.img_label.setStyleSheet("border: 2px solid #ff4d4f; background: #fff2f0;")
        
        # 查询数据库获取料号信息
        self.query_material_info(material_id)
    
    def query_material_info(self, material_id):
        """查询料号信息"""
        db_path = os.path.join(os.path.dirname(__file__), '3c.db')
        if not os.path.exists(db_path):
            return
        
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT laserid, customer, project_name, type, remark FROM cccindex WHERE laserid = ?", (material_id,))
            result = cursor.fetchone()
            conn.close()
            
            if result:
                self.selected_material = result
                laserid, customer, project_name, type_, remark = result
                print(f"料号信息: {laserid} - {customer} - {project_name} [{type_}]")
            else:
                self.selected_material = None
                print(f"未找到料号: {material_id}")
        except Exception as e:
            print(f"查询料号失败: {e}")
            self.selected_material = None

    def calculate_dimensions(self):
        """计算尺寸（参考cut.py的逻辑）"""
        try:
            order_width = float(self.target_x.text()) if self.target_x.text() else 0
            order_height = float(self.target_y.text()) if self.target_y.text() else 0
        except ValueError:
            order_width = 0
            order_height = 0
        
        edge = int(self.edge_combo.currentText())
        
        # 计算切割尺寸（目标尺寸 + 磨边量）
        cut_width = order_width + edge * 2
        cut_height = order_height + edge * 2
        
        # 计算 cutsizeX/cutsizeY (取较大/较小值)
        cutsizeX = max(cut_width, cut_height)
        cutsizeY = min(cut_width, cut_height)
        
        # 原片尺寸
        raw_width = cutsizeX + 30
        raw_height = cutsizeY + 30
        
        # 计算 displayX/displayY (取较大/较小值)
        displayX = max(order_width, order_height)
        displayY = min(order_width, order_height)
        
        return {
            'order_width': order_width,
            'order_height': order_height,
            'cut_width': cut_width,
            'cut_height': cut_height,
            'cutsizeX': cutsizeX,
            'cutsizeY': cutsizeY,
            'raw_width': raw_width,
            'raw_height': raw_height,
            'displayX': displayX,
            'displayY': displayY,
            'edge': edge,
        }

    def get_material_code(self):
        """获取料号代码"""
        if not self.selected_material:
            return ""
        
        glass_type = self.glass_type_combo.currentText()
        thickness = self.thickness_combo.currentText()
        
        # 根据类型生成料号
        if glass_type == '白玻':
            return f"{thickness}mm白玻"
        elif glass_type == '玉':
            return f"{thickness}mm玉"
        elif glass_type == 'LOWE':
            return f"{thickness}mmlowe"
        
        return f"{thickness}mm"

    def preview_gcode(self):
        """生成预览"""
        dims = self.calculate_dimensions()
        
        # 获取料号
        material_code = self.get_material_code()
        
        # 订单尺寸字符串
        order_size = f"{int(dims['order_width'])}x{int(dims['order_height'])}"
        
        # 切割原点
        cut_x = 0
        cut_y = 0
        
        # 客户名称
        customer_name = ""
        
        # 订单号/分组号
        order_number = ""
        group_number = ""
        
        # 位置信息 - 3C码位置格式：[料号 位置]
        laserid = self.selected_material[0] if self.selected_material else ""
        pos_3c = str(self.pos_3c.currentIndex() + 1)
        code_3c_position = f"{laserid} {pos_3c}"
        
        # DM码位置自动计算：X>Y时为2，X<Y时为1
        if dims['displayX'] > dims['displayY']:
            dm_code_position = "2"
        else:
            dm_code_position = "1"
        
        # DM码（使用料号作为DM码）
        dm_code = self.selected_material[0] if self.selected_material else ""
        
        # 基准边 = 输入的X值
        reference_edge = f"{int(dims['order_width'])}"
        
        # 计算切割方向
        cutYdirection = dims['raw_height'] - 3
        
        # 生成G代码
        g_code = f"""N01  P3000 = {int(dims['raw_width'])}
N02  P3001 = {int(dims['raw_height'])}
N03  P3002 = 0
N04  P3003 = 0
N05  P3004 = 0
N06  P3005 = 0
N07  P3006 = 1
N08  P3007 = {material_code}
N09  P3008 = 1
N10  P3009 = 1
N11  P3010 = 
N12  P3011 = {self.thickness_combo.currentText()}
N13  P4001= {int(dims['cutsizeX'])}_{int(dims['cutsizeY'])}_{cut_x}_{cut_y}_{int(dims['displayX'])}_{int(dims['displayY'])}_{customer_name}_1_{group_number}_{order_number}_{dm_code}____{dm_code}__{order_size}_{reference_edge}_{group_number}_{code_3c_position}_3 {dm_code_position}_________

G17
G92 X0 Y0
G90
G00 X3 Y{int(dims['cutsizeY'])}
M03
M09
G01 X{int(dims['cutsizeX'])} Y{int(dims['cutsizeY'])}
M10
G00 X{int(dims['cutsizeX'])} Y3
M09
G01 X{int(dims['cutsizeX'])} Y{int(cutYdirection)}
M10
M04
G90G00X0Y0Z0
M23
M24
M30
"""
        self.preview_edit.setText(g_code)
        self.generated_gcode = g_code

    def export_gcode(self):
        """导出G文件"""
        if not hasattr(self, 'generated_gcode') or not self.generated_gcode:
            self.preview_gcode()
        
        # 生成默认文件名
        dims = self.calculate_dimensions()
        order_size = f"{int(dims['order_width'])}x{int(dims['order_height'])}"
        order_size_clean = order_size.replace('/', '').replace('\\', '')
        default_name = f"{order_size_clean}_1_{int(dims['raw_width'])}_{int(dims['raw_height'])}.g"
        
        # 选择保存位置
        filename, _ = QFileDialog.getSaveFileName(
            self, "保存G文件", 
            os.path.join(self.GCODE_EXPORT_DIR, default_name),
            "G Files (*.g);;All Files (*)"
        )
        
        if filename:
            try:
                # 确保目录存在
                os.makedirs(os.path.dirname(filename), exist_ok=True)
                
                with open(filename, 'w', encoding='gb18030') as f:
                    f.write(self.generated_gcode)
                print(f"G文件已导出: {filename}")
            except Exception as e:
                print(f"导出失败: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = GCreteWindow()
    window.show()
    sys.exit(app.exec_())
