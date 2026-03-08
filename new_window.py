import sys
import os
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QMenuBar, QMenu, QAction, QStatusBar,
                             QLineEdit, QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt

class GCodeParser:
    def __init__(self):
        self.header_pattern = re.compile(r'N(\d+)\s+P(\d+)\s+=\s+(.+)')
        self.data_pattern = re.compile(r'N(\d+)\s+P(\d+)\s+=\s+(.+)')

    def parse_file(self, file_path):
        try:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.readlines()
            except UnicodeDecodeError:
                with open(file_path, 'r', encoding='gb2312') as f:
                    content = f.readlines()

            pieces_data = []
            common_data = {}
            
            processing_pieces = False
            current_piece = 0
            
            for line_num, line in enumerate(content, 1):
                line = line.strip()
                if not line:
                    continue

                if line.startswith('G'):
                    break

                match = self.header_pattern.match(line)
                if match:
                    n_num, p_num, value = match.groups()
                    value = value.strip()
                    
                    try:
                        if n_num == '13' and p_num.startswith('4'):
                            processing_pieces = True
                        
                        if processing_pieces and p_num.startswith('4'):
                            current_piece += 1
                            parts = value.split('_')
                            if len(parts) >= 22:
                                piece_data = common_data.copy()
                                piece_data.update({
                                    'cut_width': int(parts[0]),
                                    'cut_height': int(parts[1]),
                                    'cut_x': int(parts[2]),
                                    'cut_y': int(parts[3]),
                                    'order_width': int(parts[4]),
                                    'order_height': int(parts[5]),
                                    'customer_name': parts[6],
                                    'piece_number': current_piece,
                                    'order_number': parts[9],
                                    'dm_code': parts[10],
                                    'order_size': parts[16],
                                    'reference_edge': parts[17],
                                    'group_number': parts[18],
                                    'code_3c_position': parts[19],
                                    'dm_code_position': parts[20],
                                    'tiya_3c_position': parts[21]
                                })
                                pieces_data.append(piece_data)
                        elif not processing_pieces:
                            if p_num == '3000':
                                common_data['raw_width'] = int(value)
                            elif p_num == '3001':
                                common_data['raw_height'] = int(value)
                            elif p_num == '3007':
                                common_data['material_code'] = value
                            elif p_num == '3008':
                                common_data['layout_number'] = int(value)
                            elif p_num == '3009':
                                common_data['total_layouts'] = int(value)
                            elif p_num == '3011':
                                common_data['thickness'] = int(value)
                    except (ValueError, IndexError):
                        continue
                elif line.startswith('N13') and not line.startswith('G'):
                    continue
            return pieces_data

        except Exception:
            return None

# 程序功能：
# 1. 扫描文件夹
# 2. 搜索时遍历解析G代码并显示结果
# 3. 搜索数据
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # 设置默认工作目录
        self.work_dir = r'\\landierp\Share\激光文件'
        self.init_ui()

    def select_work_dir(self):
        # 设置默认目录为网络共享路径
        default_dir = r'\\landierp\Share\激光文件'
        dir_path = QFileDialog.getExistingDirectory(self, '选择工作目录', default_dir)
        if dir_path:
            self.work_dir = dir_path
            self.statusBar().showMessage(f'已选择工作目录: {dir_path}')

    def adjust_column_widths(self, *args):  # 将方法移到类级别
        total_width = self.table.viewport().width()
        # 设置固定的列宽比例
        self.table.setColumnWidth(0, int(total_width * 0.2))  # 文件名列 20%
        self.table.setColumnWidth(1, int(total_width * 0.15))  # 客户名列 15%
        self.table.setColumnWidth(2, int(total_width * 0.15))  # 订单尺寸列 20%
        self.table.setColumnWidth(3, int(total_width * 0.1))  # 3C料号列 10%
        self.table.setColumnWidth(4, int(total_width * 0.1))  # DM码列 10%
        self.table.setColumnWidth(5, int(total_width * 0.1))  # DM码位置列 10%
        self.table.setColumnWidth(6, int(total_width * 0.1))  # 原材料类型列 10%
        self.table.setColumnWidth(7, int(total_width * 0.1))  # 分组号列 5%
        

    def init_ui(self):
        # 设置窗口基本属性
        self.setWindowTitle('G文件尺寸搜索')
        self.setGeometry(500, 200, 1200, 700)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建主布局
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 创建搜索区域
        search_widget = QWidget()
        search_layout = QHBoxLayout(search_widget)
        search_layout.setContentsMargins(0, 0, 0, 0)

        # 工作目录选择
        self.dir_button = QPushButton('选择工作目录')
        self.dir_button.clicked.connect(self.select_work_dir)
        search_layout.addWidget(self.dir_button)

        # 搜索关键字输入框
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('请输入搜索关键字')
        self.search_input.returnPressed.connect(self.search_files)
        search_layout.addWidget(self.search_input)

        # 搜索按钮
        search_button = QPushButton('搜索')
        search_button.clicked.connect(self.search_files)
        search_layout.addWidget(search_button)

        layout.addWidget(search_widget)

        # 创建结果表格
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(['文件名', '客户名', '订单尺寸', '3C料号', 'DM码', 'DM码位置', '原材料类型', '分组号'])
        # 设置表格总宽度和列宽自适应
        self.table.horizontalHeader().setStretchLastSection(True)

        self.table.verticalHeader().setDefaultSectionSize(40)  # 设置行高
        # 设置各列宽度比例
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)  # 文件名列
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)  # 订单尺寸列
        for col in range(2, 8):
            self.table.horizontalHeader().setSectionResizeMode(col, QHeaderView.Interactive)  # 其他列
        
        # 初始调整列宽
        self.adjust_column_widths()
        # 当表格大小改变时重新调整列宽
        self.table.horizontalHeader().sectionResized.connect(self.adjust_column_widths)
        layout.addWidget(self.table)

        # 设置统一样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                padding: 10px 20px;
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 500;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #1565C0;
            }
            QLineEdit {
                padding: 10px 15px;
                border: 2px solid #E0E0E0;
                border-radius: 6px;
                font-size: 14px;
                background-color: white;
                min-width: 200px;
            }
            QLineEdit:focus {
                border-color: #2196F3;
                background-color: #F5F5F5;
            }
            QTableWidget {
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                background-color: white;
                gridline-color: #EEEEEE;
                selection-background-color: #E3F2FD;
                selection-color: #1565C0;
            }
            QHeaderView::section {
                background-color: #F5F5F5;
                padding: 12px 12px;
                border: none;
                border-right: 1px solid #E0E0E0;
                border-bottom: 2px solid #2196F3;
                font-weight: bold;
                color: #424242;
                font-size: 13px;
            }
            QTableWidget::item {
                padding: 14px;
                border-bottom: 1px solid #EEEEEE;
            }
            QTableWidget::item:selected {
                background-color: #E3F2FD;
            }
            QStatusBar {
                background-color: #ffffff;
                color: #666666;
            }
        """)

        # 添加状态栏
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        status_bar.showMessage('就绪')

    def on_new(self):
        self.statusBar().showMessage('创建新文件')

    def search_files(self):
        if not hasattr(self, 'work_dir'):
            self.statusBar().showMessage('请先选择工作目录')
            return

        keyword = self.search_input.text().strip().lower()  # 转换为小写以进行不区分大小写的搜索
        if not keyword:
            self.statusBar().showMessage('请输入搜索关键字')
            return

        self.table.setRowCount(0)
        self.statusBar().showMessage('正在搜索...')
        QApplication.processEvents()  # 保持界面响应

        try:
            parser = GCodeParser()
            found_count = 0
            
            for root, _, files in os.walk(self.work_dir):
                for file in files:
                    if file.endswith('.g'):
                        file_path = os.path.join(root, file)
                        try:
                            pieces_data = parser.parse_file(file_path)
                            if pieces_data:
                                for piece in pieces_data:
                                    # 优化搜索逻辑，只在需要的字段中搜索
                                    search_fields = [
                                        str(piece.get('customer_name', '')),
                                        str(piece.get('order_size', '')),
                                        str(piece.get('dm_code', '')),
                                        str(piece.get('code_3c_position', '')),
                                        str(piece.get('material_code', ''))
                                    ]
                                    
                                    if any(keyword in field.lower() for field in search_fields):
                                        row = self.table.rowCount()
                                        self.table.insertRow(row)
                                        
                                        # 设置单元格对齐方式
                                        items = [
                                            file,
                                            piece.get('customer_name', ''),
                                            piece.get('order_size', ''),
                                            piece.get('code_3c_position', '') or piece.get('tiya_3c_position', ''),
                                            piece.get('dm_code', ''),
                                            piece.get('dm_code_position', ''),
                                            piece.get('material_code', ''),
                                            piece.get('group_number', '')
                                        ]
                                        
                                        for col, text in enumerate(items):
                                            item = QTableWidgetItem(str(text))
                                            item.setTextAlignment(Qt.AlignCenter)
                                            self.table.setItem(row, col, item)
                                            
                                        found_count += 1
                                        if found_count % 10 == 0:  # 每10条更新一次状态
                                            self.statusBar().showMessage(f'已找到 {found_count} 个匹配项...')
                                            QApplication.processEvents()
                                            
                        except Exception as e:
                            print(f'解析文件 {file_path} 时出错: {str(e)}')

            self.statusBar().showMessage(f'搜索完成，找到 {found_count} 个匹配项')
            
        except Exception as e:
            self.statusBar().showMessage(f'搜索出错: {str(e)}')

        # 在创建表格后添加行高设置
        self.table.verticalHeader().setDefaultSectionSize(100)  # 设置默认行高为50像素
        
        # 修改表格样式中的padding和字体大小
        self.setStyleSheet("""
            QHeaderView::section {
                background-color: #F5F5F5;
                padding: 20px 12px;
                border: none;
                border-right: 1px solid #E0E0E0;
                border-bottom: 2px solid #2196F3;
                font-weight: bold;
                color: #424242;
                font-size: 14px;
            }
            QTableWidget::item {
                padding: 16px;
                border-bottom: 1px solid #EEEEEE;
                font-size: 14px;
            }
            QTableWidget::item:selected {
                background-color: #E3F2FD;
            }
            QStatusBar {
                background-color: #ffffff;
                color: #666666;
            }
        """)

        # 添加状态栏
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        status_bar.showMessage('就绪')

    def on_new(self):
        self.statusBar().showMessage('创建新文件')

    def search_files(self):
        if not hasattr(self, 'work_dir'):
            self.statusBar().showMessage('请先选择工作目录')
            return

        keyword = self.search_input.text().strip().lower()  # 转换为小写以进行不区分大小写的搜索
        if not keyword:
            self.statusBar().showMessage('请输入搜索关键字')
            return

        self.table.setRowCount(0)
        self.statusBar().showMessage('正在搜索...')
        QApplication.processEvents()  # 保持界面响应

        try:
            parser = GCodeParser()
            found_count = 0
            
            for root, _, files in os.walk(self.work_dir):
                for file in files:
                    if file.endswith('.g'):
                        file_path = os.path.join(root, file)
                        try:
                            pieces_data = parser.parse_file(file_path)
                            if pieces_data:
                                for piece in pieces_data:
                                    # 优化搜索逻辑，只在需要的字段中搜索
                                    search_fields = [
                                        str(piece.get('customer_name', '')),
                                        str(piece.get('order_size', '')),
                                        str(piece.get('dm_code', '')),
                                        str(piece.get('code_3c_position', '')),
                                        str(piece.get('material_code', ''))
                                    ]
                                    
                                    if any(keyword in field.lower() for field in search_fields):
                                        row = self.table.rowCount()
                                        self.table.insertRow(row)
                                        
                                        # 设置单元格对齐方式
                                        items = [
                                            file,
                                            piece.get('customer_name', ''),
                                            piece.get('order_size', ''),
                                            piece.get('code_3c_position', '') or piece.get('tiya_3c_position', ''),
                                            piece.get('dm_code', ''),
                                            piece.get('dm_code_position', ''),
                                            piece.get('material_code', ''),
                                            piece.get('group_number', '')
                                        ]
                                        
                                        for col, text in enumerate(items):
                                            item = QTableWidgetItem(str(text))
                                            item.setTextAlignment(Qt.AlignCenter)
                                            self.table.setItem(row, col, item)
                                            
                                        found_count += 1
                                        if found_count % 10 == 0:  # 每10条更新一次状态
                                            self.statusBar().showMessage(f'已找到 {found_count} 个匹配项...')
                                            QApplication.processEvents()
                                            
                        except Exception as e:
                            print(f'解析文件 {file_path} 时出错: {str(e)}')

            self.statusBar().showMessage(f'搜索完成，找到 {found_count} 个匹配项')
            
        except Exception as e:
            self.statusBar().showMessage(f'搜索出错: {str(e)}')

# 修改程序入口点
if __name__ == "__main__":
    # 重置系统错误输出
    import sys
    import io
    sys.stderr = sys.__stderr__
    
    # 设置较小的递归限制
    sys.setrecursionlimit(500)
    
    # 基本初始化
    app = QApplication(sys.argv)
    
    try:
        # 创建主窗口
        window = MainWindow()
        window.show()
        
        # 直接执行事件循环
        app.exec_()
        
    except Exception as e:
        print(f"错误: {e}")
        
    finally:
        # 确保程序正常退出
        app.quit()