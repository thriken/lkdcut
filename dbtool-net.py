import os
import sqlite3
import socket
import subprocess
import warnings
import time
import threading
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox, QProgressDialog,
    QTextEdit
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer


# 忽略libpng警告
warnings.filterwarnings("ignore", category=UserWarning, message="iCCP")

# 优化后的数据库管理器类 - 线程安全版本
class DatabaseManager:
    def __init__(self):
        # 定义两个基地的配置
        self.site_a = {
            'name': '新丰基地',
            'gateway': '192.168.8.1',
            'subnet': '192.168.8.0/22',  # 192.168.8.1 - 192.168.11.254
            'servername': 'LANDIERP',
            'db_path_ip': r"\\192.168.9.250\办公室\补片程序\glass_data.db",  # IP路径
            'db_path_name': r"\\LANDIERP\办公室\补片程序\glass_data.db"  # 服务器名路径
        }
        self.site_b = {
            'name': '信义基地', 
            'gateway': '192.168.100.1',
            'subnet': '192.168.100.0/24',  # 192.168.100.1 - 192.168.100.254
            'servername': 'XYERP',
            'db_path_ip': r"\\192.168.100.200\Share\glass_data.db",  # IP路径
            'db_path_name': r"\\XYERP\Share\glass_data.db"  # 服务器名路径
        }
        
        # 线程本地存储数据库连接
        self.local_storage = threading.local()
        # 检测当前网络环境并选择数据库
        self.current_site = self.detect_network_environment()
        
        # 性能监控数据
        self.performance_stats = {
            'total_searches': 0,
            'total_deletes': 0,
            'avg_search_time': 0,
            'avg_delete_time': 0,
            'last_operation_time': 0
        }
        
    def get_thread_db_connection(self):
        """获取当前线程的数据库连接"""
        if not hasattr(self.local_storage, 'connection'):
            # 创建新的连接，设置超时和重试
            self.local_storage.connection = self.connect_with_retry()
        return self.local_storage.connection
    
    def connect_with_retry(self):
        """带重试机制的连接方法"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                return self.connect()
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                print(f"连接尝试 {attempt + 1} 失败: {e}")
                time.sleep(1)
        return None

    def detect_network_environment(self):
        """检测当前网络环境，返回对应的基地配置"""
        try:
            # 获取本机IP地址
            hostname = socket.gethostname()
            local_ip = socket.gethostbyname(hostname)
            
            print(f"当前主机IP: {local_ip}")
            
            # 检查新丰基地网络段 (192.168.8.0/22)
            if self.is_ip_in_subnet(local_ip, self.site_a['subnet']):
                print(f"检测到新丰基地网络环境")
                return self.site_a
            
            # 检查信义基地网络段 (192.168.100.0/24)
            elif self.is_ip_in_subnet(local_ip, self.site_b['subnet']):
                print(f"检测到信义基地网络环境")
                return self.site_b
            
            # 尝试ping网关检测网络连通性
            print("正在通过网关检测网络环境...")
            if self.ping_gateway(self.site_a['gateway']):
                print(f"通过网关检测到新丰基地网络")
                return self.site_a
            elif self.ping_gateway(self.site_b['gateway']):
                print(f"通过网关检测到信义基地网络")
                return self.site_b
            
            print("无法确定网络环境")
            return None
            
        except Exception as e:
            print(f"网络环境检测失败: {e}")
            return None
    
    def is_ip_in_subnet(self, ip, subnet):
        """检查IP是否在指定子网内"""
        try:
            ip_parts = list(map(int, ip.split('.')))
            
            if '/' in subnet:
                network_str, mask_bits = subnet.split('/')
                mask_bits = int(mask_bits)
                network_parts = list(map(int, network_str.split('.')))
                
                # 计算子网掩码
                mask = (0xffffffff << (32 - mask_bits)) & 0xffffffff
                
                # 将IP和网络地址转换为整数
                ip_int = (ip_parts[0] << 24) + (ip_parts[1] << 16) + (ip_parts[2] << 8) + ip_parts[3]
                network_int = (network_parts[0] << 24) + (network_parts[1] << 16) + (network_parts[2] << 8) + network_parts[3]
                
                # 检查是否在同一个子网
                return (ip_int & mask) == (network_int & mask)
            
        except Exception as e:
            print(f"子网检测错误: {e}")
            return False
    
    def ping_gateway(self, gateway):
        """尝试ping网关检测网络连通性"""
        try:
            # 使用ping命令检测网关连通性
            result = subprocess.run(['ping', '-n', '1', '-w', '1000', gateway], 
                                  capture_output=True, text=True, timeout=2)
            return result.returncode == 0
        except:
            return False
    
    def try_connect_site(self, site):
        """尝试连接指定基地的数据库，先试IP后试服务器名"""
        connection_method = ""
        
        print(f"尝试连接 {site['name']} 数据库...")
        
        # 先尝试IP路径
        print(f"尝试IP路径: {site['db_path_ip']}")
        if os.path.exists(site['db_path_ip']):
            try:
                connection_method = "IP地址"
                conn = sqlite3.connect(site['db_path_ip'])
                print(f"通过IP地址连接成功: {site['db_path_ip']}")
                return conn, connection_method
            except Exception as e:
                print(f"IP地址连接失败: {e}")
        
        # 如果IP路径失败，尝试服务器名路径
        print(f"尝试服务器名路径: {site['db_path_name']}")
        if os.path.exists(site['db_path_name']):
            try:
                connection_method = "服务器名"
                conn = sqlite3.connect(site['db_path_name'])
                print(f"通过服务器名连接成功: {site['db_path_name']}")
                return conn, connection_method
            except Exception as e:
                print(f"服务器名连接失败: {e}")
        
        print(f"{site['name']} 数据库连接失败")
        return None, None
    
    def connect(self):
        """连接到当前网络环境对应的数据库"""
        connection_info = ""
        
        # 如果检测到当前网络环境，优先连接该基地
        if self.current_site:
            conn, method = self.try_connect_site(self.current_site)
            if conn:
                connection_info = f"连接到{self.current_site['name']}的数据库 (通过{method})"
                print(connection_info)
                self.connection_info = connection_info
                return conn
        
        # 如果首选连接失败，尝试连接两个基地作为备用
        print("首选连接失败，尝试备用连接")
        
        # 尝试新丰基地
        conn, method = self.try_connect_site(self.site_a)
        if conn:
            connection_info = f"连接到新丰基地的数据库 (通过{method})"
            print(connection_info)
            self.connection_info = connection_info
            return conn
        
        # 尝试信义基地
        conn, method = self.try_connect_site(self.site_b)
        if conn:
            connection_info = f"连接到信义基地的数据库 (通过{method})"
            print(connection_info)
            self.connection_info = connection_info
            return conn
        
        # 如果所有连接都失败，弹出错误提示
        error_msg = "无法连接到数据库！\n\n"
        error_msg += "请检查以下问题：\n"
        error_msg += "1. 网络连接是否正常\n"
        error_msg += "2. 共享文件夹是否已正确设置\n"
        error_msg += "3. 数据库文件路径是否正确\n"
        error_msg += "4. 防火墙设置是否允许访问共享\n\n"
        error_msg += "需要修复共享问题后重新启动程序。"
        
        raise Exception(error_msg)

    def search_files_by_keyword(self, keyword):
        """
        根据文件名关键字搜索文件
        :param keyword: 文件名关键字
        :return: 文件信息列表 [(id, directory, file_name), ...]
        """
        try:
            connection = self.get_thread_db_connection()
            cursor = connection.cursor()
            cursor.execute('''
                SELECT id, directory, file_name 
                FROM file_info 
                WHERE file_name LIKE LOWER(?) 
                ORDER BY directory, file_name
            ''', (f'%{keyword}%',))
            results = cursor.fetchall()
            cursor.close()
            return results
        except Exception as e:
            print(f"搜索文件时出错: {e}")
            return []

    def get_glass_data_count_by_file_ids(self, file_ids):
        """
        获取指定file_id列表在glass_data表中的总记录数量
        :param file_ids: 文件ID列表
        :return: 总记录数量
        """
        try:
            if not file_ids:
                return 0
                
            connection = self.get_thread_db_connection()
            cursor = connection.cursor()
            
            # 使用 IN 子句查询多个file_id的记录数
            placeholders = ','.join(['?'] * len(file_ids))
            cursor.execute(f'''
                SELECT COUNT(*) 
                FROM glass_data 
                WHERE file_id IN ({placeholders})
            ''', file_ids)
            result = cursor.fetchone()
            cursor.close()
            return result[0] if result else 0
        except Exception as e:
            print(f"获取glass_data总数时出错: {e}")
            return 0

    def get_glass_data_count_by_file_id(self, file_id):
        """
        获取指定file_id在glass_data表中的记录数量
        :param file_id: 文件ID
        :return: 记录数量
        """
        try:
            connection = self.get_thread_db_connection()
            cursor = connection.cursor()
            cursor.execute('''
                SELECT COUNT(*) 
                FROM glass_data 
                WHERE file_id = ?
            ''', (file_id,))
            result = cursor.fetchone()
            cursor.close()
            return result[0] if result else 0
        except Exception as e:
            print(f"获取glass_data数量时出错: {e}")
            return 0

    def delete_data_by_file_ids(self, file_ids):
        """
        根据文件ID列表删除数据
        :param file_ids: 文件ID列表
        :return: 删除结果字典 {'file_info_count': x, 'glass_data_count': y}
        """
        try:
            if not file_ids:
                return {'file_info_count': 0, 'glass_data_count': 0}
            
            connection = self.get_thread_db_connection()
            cursor = connection.cursor()
            
            # 先删除glass_data表中的数据
            glass_data_count = 0
            for file_id in file_ids:
                cursor.execute('''
                    DELETE FROM glass_data 
                    WHERE file_id = ?
                ''', (file_id,))
                glass_data_count += cursor.rowcount
            
            # 再删除file_info表中的数据
            file_info_count = 0
            for file_id in file_ids:
                cursor.execute('''
                    DELETE FROM file_info 
                    WHERE id = ?
                ''', (file_id,))
                file_info_count += cursor.rowcount
            
            connection.commit()
            cursor.close()
            
            return {
                'file_info_count': file_info_count,
                'glass_data_count': glass_data_count
            }
        except Exception as e:
            print(f"删除数据时出错: {e}")
            raise

# 搜索线程类
class SearchThread(QThread):
    search_finished = pyqtSignal(list, int, str, int, dict, float)
    search_error = pyqtSignal(str)
    
    def __init__(self, db_manager, keyword):
        super().__init__()
        self.db_manager = db_manager
        self.keyword = keyword
    
    def run(self):
        start_time = time.time()
        try:
            # 创建线程特定的数据库连接
            connection = self.db_manager.connect_with_retry()
            
            # 搜索文件
            cursor = connection.cursor()
            cursor.execute('''
                SELECT id, directory, file_name 
                FROM file_info 
                WHERE file_name LIKE LOWER(?) 
                ORDER BY directory, file_name
            ''', (f'%{self.keyword}%',))
            results = cursor.fetchall()
            cursor.close()
            
            # 计算每个文件的玻璃数据数量
            file_glass_counts = {}
            for file_id, directory, file_name in results:
                cursor = connection.cursor()
                cursor.execute('''
                    SELECT COUNT(*) 
                    FROM glass_data 
                    WHERE file_id = ?
                ''', (file_id,))
                result = cursor.fetchone()
                file_glass_counts[file_id] = result[0] if result else 0
                cursor.close()
            
            # 获取总玻璃数据数量
            file_ids = [file_id for file_id, _, _ in results]
            total_glass_count = 0
            if file_ids:
                cursor = connection.cursor()
                placeholders = ','.join(['?'] * len(file_ids))
                cursor.execute(f'''
                    SELECT COUNT(*) 
                    FROM glass_data 
                    WHERE file_id IN ({placeholders})
                ''', file_ids)
                result = cursor.fetchone()
                total_glass_count = result[0] if result else 0
                cursor.close()
            
            connection.close()
            
            # 计算实际耗时
            execution_time = time.time() - start_time
            
            # 发送完成信号（包含实际耗时）
            self.search_finished.emit(results, len(results), self.keyword, total_glass_count, file_glass_counts, execution_time)
                
        except Exception as e:
            execution_time = time.time() - start_time
            self.search_error.emit(str(e))

# 删除线程类
class DeleteThread(QThread):
    delete_finished = pyqtSignal(dict, str)
    delete_error = pyqtSignal(str)
    
    def __init__(self, db_manager, file_ids, keyword):
        super().__init__()
        self.db_manager = db_manager
        self.file_ids = file_ids
        self.keyword = keyword
    
    def run(self):
        try:
            # 创建线程特定的数据库连接
            connection = self.db_manager.connect_with_retry()
            cursor = connection.cursor()
            
            # 执行删除操作
            glass_data_count = 0
            for file_id in self.file_ids:
                cursor.execute('''
                    DELETE FROM glass_data 
                    WHERE file_id = ?
                ''', (file_id,))
                glass_data_count += cursor.rowcount
            
            file_info_count = 0
            for file_id in self.file_ids:
                cursor.execute('''
                    DELETE FROM file_info 
                    WHERE id = ?
                ''', (file_id,))
                file_info_count += cursor.rowcount
            
            connection.commit()
            cursor.close()
            connection.close()
            
            result = {
                'file_info_count': file_info_count,
                'glass_data_count': glass_data_count
            }
            self.delete_finished.emit(result, self.keyword)
                
        except Exception as e:
            self.delete_error.emit(str(e))

# 优化后的主窗口类
class DBToolWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self.current_results = None  # 存储当前搜索结果
        self.current_file_glass_counts = {}  # 存储每个文件的玻璃数据数量
        
        # 线程管理
        self.search_thread = None
        self.delete_thread = None
        
        self.init_ui()
        
        # 启动定时器进行性能监控
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.update_status)
        self.status_timer.start(2000)  # 每2秒更新一次状态

    def init_ui(self):
        # 设置窗口基本属性
        title = '数据库删除工具（优化网络版）'
        if self.db.current_site:
            title += f" - {self.db.current_site['name']}"
        self.setWindowTitle(title)
        self.setGeometry(400, 300, 800, 500)
        
        # 创建主布局
        main_layout = QVBoxLayout()
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
        # 创建状态显示区域
        status_layout = QHBoxLayout()
        self.status_label = QLabel('正在初始化...')
        self.status_label.setStyleSheet("""
            QLabel {
                color: #666;
                font-size: 12px;
                padding: 8px;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                background-color: #f9f9f9;
            }
        """)
        status_layout.addWidget(self.status_label)
        main_layout.addLayout(status_layout)
        
        # 创建搜索区域
        search_layout = QHBoxLayout()
        
        keyword_label = QLabel('加工号:')
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText('请输入加工号关键字（支持模糊搜索），如K1234')
        
        # 设置自动大写转换
        self.keyword_input.textChanged.connect(self.auto_uppercase)
        
        search_button = QPushButton('搜索（线程安全）')
        search_button.clicked.connect(self.search_data_threaded)
        
        delete_button = QPushButton('删除全部（线程安全）')
        delete_button.clicked.connect(self.delete_all_threaded)
        
        search_layout.addWidget(keyword_label)
        search_layout.addWidget(self.keyword_input)
        search_layout.addWidget(search_button)
        search_layout.addWidget(delete_button)
        
        main_layout.addLayout(search_layout)
        
        # 创建结果显示区域
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                padding: 10px;
                font-family: "Consolas", "Monaco", monospace;
                font-size: 14px;
            }
        """)
        main_layout.addWidget(self.result_text)
        
        # 初始显示
        initial_text = "数据库删除工具（优化网络版）\n"
        initial_text += "================================\n"
        initial_text += "特性：\n"
        initial_text += "- 线程安全的数据库操作\n"
        initial_text += "- 性能监控和统计\n"
        initial_text += "- 自动网络环境检测\n"
        initial_text += "- 重试机制和错误处理\n\n"
        
        if hasattr(self.db, 'connection_info'):
            initial_text += f"{self.db.connection_info}\n"
        elif self.db.current_site:
            initial_text += f"当前网络环境: {self.db.current_site['name']}\n"
        else:
            initial_text += "当前网络环境: 无法确定\n"
        
        initial_text += "\n请输入关键字并点击【搜索】按钮开始搜索"
        self.result_text.setText(initial_text)
        
        # 连接回车键搜索
        self.keyword_input.returnPressed.connect(self.search_data_threaded)

    def update_status(self):
        """更新状态显示"""
        stats = self.db.performance_stats
        status_text = f"搜索统计: 总次数 {stats['total_searches']} | 平均耗时 {stats['avg_search_time']:.3f}s"
        
        if hasattr(self.db, 'connection_info'):
            status_text += f" | {self.db.connection_info}"
        elif self.db.current_site:
            status_text += f" | 当前网络: {self.db.current_site['name']}"
            
        self.status_label.setText(status_text)

    def showEvent(self, event):
        """窗口显示时自动聚焦到输入框"""
        super().showEvent(event)
        self.keyword_input.setFocus()
        
    def auto_uppercase(self, text):
        """自动将输入转换为大写"""
        cursor_position = self.keyword_input.cursorPosition()
        uppercase_text = text.upper()
        if text != uppercase_text:
            self.keyword_input.blockSignals(True)  # 暂时阻止信号循环
            self.keyword_input.setText(uppercase_text)
            self.keyword_input.setCursorPosition(cursor_position)
            self.keyword_input.blockSignals(False)

    def search_data_threaded(self):
        """多线程搜索数据"""
        keyword = self.keyword_input.text().strip().upper()
        if not keyword:
            QMessageBox.warning(self, '警告', '请输入文件名关键字')
            return
        
        # 禁用搜索按钮，防止重复点击
        self.set_ui_enabled(False)
        
        # 显示搜索中状态
        self.result_text.setText(f"正在搜索关键字: {keyword}...\n请稍候...")
        
        # 启动搜索线程
        self.search_thread = SearchThread(self.db, keyword)
        self.search_thread.search_finished.connect(self.on_search_finished)
        self.search_thread.search_error.connect(self.on_search_error)
        self.search_thread.start()

    def on_search_finished(self, results, result_count, keyword, total_glass_count, file_glass_counts, execution_time):
        """搜索完成回调"""
        # 存储搜索结果
        self.current_results = results
        self.current_file_glass_counts = file_glass_counts
        
        # 更新性能统计
        self.db.performance_stats['total_searches'] += 1
        self.db.performance_stats['last_operation_time'] = execution_time
        if self.db.performance_stats['total_searches'] > 0:
            self.db.performance_stats['avg_search_time'] = (
                (self.db.performance_stats['avg_search_time'] * (self.db.performance_stats['total_searches'] - 1) + execution_time) / 
                self.db.performance_stats['total_searches']
            )
        
        # 生成结果显示文本
        result_text = f"=== 搜索结果 ===\n"
        result_text += f"搜索关键字: {keyword}\n"
        result_text += f"匹配原片数: {result_count} 张\n"
        result_text += f"玻璃总小片数: {total_glass_count} 片\n"
        result_text += f"搜索耗时: {execution_time:.3f}秒\n\n"
        
        if results:
            result_text += f"匹配的文件列表:\n"
            for i, (file_id, directory, file_name) in enumerate(results, 1):
                single_count = file_glass_counts.get(file_id, 0)
                result_text += f"{i:2d}. ID:{file_id:4d}|目录:{directory or '根目录':10s}|文件名:{file_name:25s}|关联小片:{single_count:4d}条\n"
        else:
            result_text += "没有找到匹配的文件"
        
        # 更新结果显示区域
        self.result_text.setText(result_text)
        
        # 启用UI
        self.set_ui_enabled(True)
        
        # 清理线程
        self.search_thread = None

    def on_search_error(self, error_msg):
        """搜索错误回调"""
        QMessageBox.critical(self, '搜索错误', f'搜索时出错: {error_msg}')
        self.result_text.setText("搜索失败，请检查网络连接和数据库状态")
        
        # 启用UI
        self.set_ui_enabled(True)
        
        # 清理线程
        self.search_thread = None

    def delete_all_threaded(self):
        """多线程删除数据"""
        if self.current_results is None:
            QMessageBox.warning(self, '警告', '请先执行搜索操作')
            return
        
        if not self.current_results:
            QMessageBox.information(self, '提示', '没有数据可删除')
            return
        
        # 获取所有文件ID
        file_ids = [file_id for file_id, _, _ in self.current_results]
        
        # 确认删除
        keyword = self.keyword_input.text().strip()
        reply = QMessageBox.question(
            self, 
            '确认删除', 
            f'确定要删除关键字 "{keyword}" 的全部数据吗？\n\n'
            f'匹配文件数量: {len(self.current_results)} 个\n'
            f'glass_data表中关联数据: {sum(self.current_file_glass_counts.values())} 条\n\n'
            f'此操作不可撤销！',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # 禁用删除按钮，防止重复点击
        self.set_ui_enabled(False)
        
        # 显示删除中状态
        self.result_text.setText("正在删除数据...\n请勿关闭窗口...")
        
        # 启动删除线程
        self.delete_thread = DeleteThread(self.db, file_ids, keyword)
        self.delete_thread.delete_finished.connect(self.on_delete_finished)
        self.delete_thread.delete_error.connect(self.on_delete_error)
        self.delete_thread.start()

    def on_delete_finished(self, result, keyword):
        """删除完成回调"""
        # 显示删除结果
        QMessageBox.information(
            self,
            '删除完成',
            f"删除操作完成！\n\n"
            f"删除文件记录数: {result['file_info_count']}\n"
            f"删除玻璃数据记录数: {result['glass_data_count']}\n\n"
            f"所有与关键字 '{keyword}' 相关的数据已全部删除"
        )
        
        # 清空搜索结果
        self.current_results = None
        self.current_file_glass_counts = {}
        self.result_text.setText("删除完成，请重新搜索其他关键字")
        self.keyword_input.clear()
        
        # 启用UI
        self.set_ui_enabled(True)
        
        # 清理线程
        self.delete_thread = None

    def on_delete_error(self, error_msg):
        """删除错误回调"""
        QMessageBox.critical(self, '删除错误', f'删除时出错: {error_msg}')
        self.result_text.setText("删除失败，请检查网络连接和数据库状态")
        
        # 启用UI
        self.set_ui_enabled(True)
        
        # 清理线程
        self.delete_thread = None

    def set_ui_enabled(self, enabled):
        """设置UI元素启用状态"""
        # 获取所有按钮并设置启用状态
        for widget in self.findChildren(QPushButton):
            widget.setEnabled(enabled)
        self.keyword_input.setEnabled(enabled)

# 主程序入口
def main():
    app = QApplication([])
    
    # 设置应用程序样式
    app.setStyleSheet("""
        QMainWindow {
            background-color: #f5f5f5;
            font-family: "Microsoft YaHei UI";
        }
        QLabel {
            color: #262626;
            font-size: 14px;
            font-weight: 500;
        }
        QPushButton {
            min-width: 100px;
            padding: 8px 15px;
            background-color: #1890ff;
            color: white;
            border: none;
            border-radius: 4px;
            font-weight: 500;
        }
        QPushButton:hover {
            background-color: #40a9ff;
        }
        QPushButton:pressed {
            background-color: #096dd9;
        }
        QPushButton:disabled {
            background-color: #d9d9d9;
            color: #8c8c8c;
        }
        QLineEdit {
            padding: 8px;
            border: 1px solid #d9d9d9;
            border-radius: 4px;
            background: #fafafa;
            min-width: 200px;
        }
        QLineEdit:focus {
            border-color: #40a9ff;
        }
    """)
    
    window = DBToolWindow()
    window.show()
    
    app.exec_()

if __name__ == '__main__':
    main()