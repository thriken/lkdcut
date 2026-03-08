import os
import sqlite3
import threading
import time
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox, QProgressDialog,
    QTextEdit
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QElapsedTimer

# === 搜索线程类 ===
class SearchThread(QThread):
    search_finished = pyqtSignal(list, int, str, int, dict)  # 添加执行时间参数和文件计数字典
    search_error = pyqtSignal(str)
    
    def __init__(self, db_manager, keyword):
        super().__init__()
        self.db_manager = db_manager
        self.keyword = keyword
        self.timer = QElapsedTimer()
    
    def run(self):
        self.timer.start()
        try:
            # 在线程中创建新的数据库管理器实例，确保线程安全
            thread_db_manager = AdvancedDatabaseManager()
            
            results = thread_db_manager.search_files_by_keyword(self.keyword)
            file_ids = [file_id for file_id, _, _ in results]
            total_glass_data_count = thread_db_manager.get_glass_data_count_by_file_ids(file_ids)
            
            # 计算每个文件的关联小片数量
            file_glass_counts = {}
            for file_id, directory, file_name in results:
                count = thread_db_manager.get_glass_data_count_by_file_id(file_id)
                file_glass_counts[file_id] = count
            
            execution_time = self.timer.elapsed()
            
            self.search_finished.emit(results, total_glass_data_count, self.keyword, execution_time, file_glass_counts)
        except Exception as e:
            self.search_error.emit(str(e))

# === 删除线程类 ===
class DeleteThread(QThread):
    delete_finished = pyqtSignal(dict, str, int)  # 添加执行时间参数
    delete_error = pyqtSignal(str)
    
    def __init__(self, db_manager, file_ids):
        super().__init__()
        self.db_manager = db_manager
        self.file_ids = file_ids
        self.timer = QElapsedTimer()
    
    def run(self):
        self.timer.start()
        try:
            # 在线程中创建新的数据库管理器实例，确保线程安全
            thread_db_manager = AdvancedDatabaseManager()
            
            result = thread_db_manager.delete_data_by_file_ids(self.file_ids)
            execution_time = self.timer.elapsed()
            self.delete_finished.emit(result, "删除完成", execution_time)
        except Exception as e:
            self.delete_error.emit(str(e))

# === 简单缓存类 ===
class SimpleCache:
    def __init__(self, max_size=100, ttl=300):
        self.max_size = max_size
        self.ttl = ttl
        self._cache = {}
        self._lock = threading.Lock()
    
    def get(self, key):
        with self._lock:
            if key in self._cache:
                data, timestamp = self._cache[key]
                if time.time() - timestamp < self.ttl:
                    return data
                else:
                    del self._cache[key]
            return None
    
    def set(self, key, value):
        with self._lock:
            if len(self._cache) >= self.max_size:
                # 移除最旧的缓存项
                oldest_key = min(self._cache.keys(), key=lambda k: self._cache[k][1])
                del self._cache[oldest_key]
            self._cache[key] = (value, time.time())

# === 数据库装饰器 ===
def db_connection(func):
    def wrapper(self, *args, **kwargs):
        # 每次调用都创建新的数据库连接，确保线程安全
        connection = sqlite3.connect(self.local_db_path)
        cursor = connection.cursor()
        try:
            result = func(self, cursor, *args, **kwargs)
            connection.commit()
            return result
        except Exception as e:
            connection.rollback()
            print(f"数据库操作出错: {e}")
            raise
        finally:
            cursor.close()
            connection.close()
    return wrapper

# === 高级数据库管理器 ===
class AdvancedDatabaseManager:
    def __init__(self):
        self.local_db_path = "glass_data.db"
        
        # 缓存系统
        self.search_cache = SimpleCache(max_size=200, ttl=600)  # 200项缓存，10分钟过期
        self.count_cache = SimpleCache(max_size=500, ttl=300)   # 500项缓存，5分钟过期
        
        # 性能统计
        self.performance_stats = {
            'search_count': 0,
            'delete_count': 0,
            'total_search_time': 0,
            'total_delete_time': 0
        }
        
        self.check_database_file()
    
    def check_database_file(self):
        if not os.path.exists(self.local_db_path):
            error_msg = f"数据库文件不存在！\n\n"
            error_msg += f"请在当前目录放置glass_data.db文件：\n"
            error_msg += f"{os.path.abspath(self.local_db_path)}"
            raise FileNotFoundError(error_msg)

    @db_connection
    def search_files_by_keyword(self, cursor, keyword):
        """搜索文件方法"""
        self.performance_stats['search_count'] += 1
        
        # 检查缓存
        cache_key = keyword.lower()
        cached_result = self.search_cache.get(cache_key)
        if cached_result:
            return cached_result
        
        # 优化搜索查询
        cursor.execute('''
            SELECT id, directory, file_name 
            FROM file_info 
            WHERE file_name LIKE ? COLLATE NOCASE
            ORDER BY directory, file_name
        ''', (f'%{keyword}%',))
        results = cursor.fetchall()
        
        # 缓存结果
        if len(results) <= 100:
            self.search_cache.set(cache_key, results)
        
        return results

    @db_connection
    def get_glass_data_count_by_file_ids(self, cursor, file_ids):
        """获取多个文件的小片总数"""
        if not file_ids:
            return 0
        
        # 检查缓存
        cache_key = f"count_{hash(tuple(file_ids))}"
        cached_result = self.count_cache.get(cache_key)
        if cached_result:
            return cached_result
        
        # 使用 IN 子句查询
        placeholders = ','.join(['?'] * len(file_ids))
        cursor.execute(f'''
            SELECT COUNT(*) 
            FROM glass_data 
            WHERE file_id IN ({placeholders})
        ''', file_ids)
        result = cursor.fetchone()
        count = result[0] if result else 0
        
        # 缓存结果
        self.count_cache.set(cache_key, count)
        
        return count

    @db_connection
    def get_glass_data_count_by_file_id(self, cursor, file_id):
        """获取单个文件的小片数量"""
        # 检查缓存
        cache_key = f"single_count_{file_id}"
        cached_result = self.count_cache.get(cache_key)
        if cached_result:
            return cached_result
        
        cursor.execute('''
            SELECT COUNT(*) 
            FROM glass_data 
            WHERE file_id = ?
        ''', (file_id,))
        result = cursor.fetchone()
        count = result[0] if result else 0
        
        # 缓存结果
        self.count_cache.set(cache_key, count)
        
        return count

    @db_connection
    def delete_data_by_file_ids(self, cursor, file_ids):
        """删除指定文件ID的数据"""
        self.performance_stats['delete_count'] += 1
        
        if not file_ids:
            return {'deleted_files': 0, 'deleted_glass_data': 0}
        
        # 先获取删除前的计数
        placeholders = ','.join(['?'] * len(file_ids))
        cursor.execute(f'''
            SELECT COUNT(*) 
            FROM glass_data 
            WHERE file_id IN ({placeholders})
        ''', file_ids)
        glass_data_count = cursor.fetchone()[0]
        
        # 删除glass_data表中的数据
        cursor.execute(f'''
            DELETE FROM glass_data 
            WHERE file_id IN ({placeholders})
        ''', file_ids)
        
        # 删除file_info表中的数据
        cursor.execute(f'''
            DELETE FROM file_info 
            WHERE id IN ({placeholders})
        ''', file_ids)
        
        # 清空相关缓存
        self.search_cache = SimpleCache(max_size=200, ttl=600)
        self.count_cache = SimpleCache(max_size=500, ttl=300)
        
        return {
            'deleted_files': len(file_ids),
            'deleted_glass_data': glass_data_count
        }

# === 主窗口类 ===
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # 当前搜索结果
        self.current_results = []
        
        # 搜索和删除线程
        self.search_thread = None
        self.delete_thread = None
        
        # 延迟初始化数据库管理器，避免启动时自动操作
        self.db = None
        
        self.init_ui()
        self.setup_status_bar()
        
        # 延迟数据库初始化，确保UI完全初始化后再连接数据库
        QTimer.singleShot(100, self.init_database)
    
    def init_database(self):
        """延迟初始化数据库，避免启动时自动搜索"""
        try:
            self.db = AdvancedDatabaseManager()
            # 确保进度条在初始化时隐藏
            if hasattr(self, 'progress'):
                self.progress.hide()
            # 直接设置状态，避免定时器干扰
            self.status_label.setText("数据库连接成功 - 就绪")
        except Exception as e:
            # 确保进度条在错误时也隐藏
            if hasattr(self, 'progress'):
                self.progress.hide()
            self.status_label.setText(f"数据库连接失败: {str(e)}")
            QMessageBox.critical(self, "数据库错误", f"无法连接数据库: {str(e)}")
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("玻璃切割数据管理工具 - 本地优化版")
        self.setGeometry(400, 300, 800, 500)
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)

        # 搜索区域
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("加工号:"))
        
        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("请输入加工号关键字（支持模糊搜索），如K1234")
        self.keyword_input.returnPressed.connect(self.perform_search)
        # 设置自动大写转换
        self.keyword_input.textChanged.connect(self.auto_uppercase)
        search_layout.addWidget(self.keyword_input)
        
        self.search_button = QPushButton("搜索")
        self.search_button.clicked.connect(self.perform_search)
        search_layout.addWidget(self.search_button)
        
        # 删除按钮
        self.delete_button = QPushButton("删除选中数据")
        self.delete_button.clicked.connect(self.confirm_delete)
        self.delete_button.setEnabled(False)
        search_layout.addWidget(self.delete_button)
        
        main_layout.addLayout(search_layout)
        
        # 结果显示区域
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
        # 设置初始显示文本
        initial_text = "玻璃切割数据管理工具（本地优化版）\n"
        initial_text += "================================\n"
        initial_text += "数据库文件：glass_data.db\n"
        initial_text += "\n请输入关键字并点击【搜索】按钮开始搜索"
        self.result_text.setText(initial_text)
        main_layout.addWidget(self.result_text)
        
        # 进度条不在初始化时创建，只在需要时创建
    
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
    
    def setup_status_bar(self):
        """设置状态栏"""
        self.status_label = QLabel("就绪")
        self.statusBar().addWidget(self.status_label)
        
        # 定时更新状态栏（但不要自动触发搜索）
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.update_status)
        self.status_timer.start(1000)  # 每秒更新一次
    
    def update_status(self):
        """更新状态栏"""
        if self.db is None:
            # 只在未初始化时显示一次，避免重复显示
            if self.status_label.text() != "数据库连接成功":
                self.status_label.setText("数据库初始化中...")
            return
            
        stats = self.db.performance_stats
        status_text = f"搜索次数: {stats['search_count']} | 删除次数: {stats['delete_count']} | "
        status_text += f"搜索总耗时: {stats['total_search_time']}ms | 删除总耗时: {stats['total_delete_time']}ms"
        self.status_label.setText(status_text)
    
    def set_ui_enabled(self, enabled):
        """设置UI控件启用状态"""
        self.search_button.setEnabled(enabled)
        self.delete_button.setEnabled(enabled and len(self.current_results) > 0)
        self.keyword_input.setEnabled(enabled)
    
    def perform_search(self):
        """执行搜索操作（仅在用户点击搜索按钮时执行）"""
        if self.db is None:
            QMessageBox.warning(self, "警告", "数据库未初始化完成，请稍后再试！")
            return
            
        keyword = self.keyword_input.text().strip()
        if not keyword:
            QMessageBox.warning(self, "警告", "请输入搜索关键字！")
            return
        
        # 禁用UI控件
        self.set_ui_enabled(False)
        
        # 创建进度条（只在搜索时创建）
        self.progress = QProgressDialog("", None, 0, 0, self)
        self.progress.setWindowModality(Qt.WindowModal)
        self.progress.setLabelText("正在搜索...")
        self.progress.setCancelButton(None)  # 确保没有取消按钮
        
        # 创建搜索线程
        self.search_thread = SearchThread(self.db, keyword)
        self.search_thread.search_finished.connect(self.on_search_finished)
        self.search_thread.search_error.connect(self.on_search_error)
        self.search_thread.start()
        
        # 显示进度条
        self.progress.show()
    
    # 删除cancel_search方法，因为不再需要取消按钮
    
    def on_search_finished(self, results, total_count, keyword, execution_time, file_glass_counts):
        """搜索完成回调"""
        self.current_results = results
        
        # 更新性能统计
        self.db.performance_stats['total_search_time'] += execution_time
        
        # 启用UI控件
        self.set_ui_enabled(True)
        
        # 隐藏进度条
        self.progress.close()
        
        result_text = f"=== 搜索结果 (耗时: {execution_time}ms) ===\n"
        result_text += f"搜索关键字: {keyword}\n"
        result_text += f"匹配原片数: {len(results)} 张\n"
        result_text += f"玻璃总小片数: {total_count} 片\n\n"
        
        if results:
            result_text += f"匹配的文件列表:\n"
            for i, (file_id, directory, file_name) in enumerate(results, 1):
                # 使用搜索线程计算好的文件关联小片数量
                single_count = file_glass_counts.get(file_id, 0)
                result_text += f"{i:2d}. ID:{file_id:4d}|目录:{directory or '根目录':10s}|文件名:{file_name:30s}|关联小片:{single_count:4d}条\n"
        else:
            result_text += "没有找到匹配的文件"
        
        self.result_text.setText(result_text)
        self.progress.close()
        self.set_ui_enabled(True)
        self.status_label.setText(f"搜索完成: 找到 {len(results)} 个文件")
    
    def on_search_error(self, error_msg):
        """搜索错误回调"""
        QMessageBox.critical(self, "错误", f"搜索时出错: {error_msg}")
        self.progress.close()
        self.set_ui_enabled(True)
    
    def confirm_delete(self):
        """确认删除操作"""
        if not self.current_results:
            QMessageBox.warning(self, "警告", "没有可删除的数据！")
            return
        
        # 获取文件ID列表
        file_ids = [file_id for file_id, _, _ in self.current_results]
        keyword = self.keyword_input.text().strip()
        
        # 确认删除
        reply = QMessageBox.question(
            self, 
            '确认删除', 
            f'确定要删除关键字 "{keyword}" 的全部数据吗？\n\n'
            f'匹配文件数量: {len(self.current_results)} 个\n'
            f'glass_data表中关联数据: 即将计算... 条\n\n'
            f'此操作不可撤销！',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.perform_delete(file_ids)
    
    def perform_delete(self, file_ids):
        """执行删除操作"""
        # 禁用UI控件
        self.set_ui_enabled(False)
        
        # 创建删除线程
        self.delete_thread = DeleteThread(self.db, file_ids)
        self.delete_thread.delete_finished.connect(self.on_delete_finished)
        self.delete_thread.delete_error.connect(self.on_delete_error)
        self.delete_thread.start()
        
        # 创建进度条（只在删除时创建）
        self.progress = QProgressDialog("", None, 0, 0, self)
        self.progress.setWindowModality(Qt.WindowModal)
        self.progress.setLabelText("正在删除数据...")
        self.progress.setCancelButton(None)  # 确保没有取消按钮
        
        # 显示进度条
        self.progress.show()
    
    def on_delete_finished(self, result, message, execution_time):
        """删除完成回调"""
        # 更新性能统计
        self.db.performance_stats['total_delete_time'] += execution_time
        
        QMessageBox.information(self, "完成", 
            f"删除操作完成！\n\n"
            f"删除文件数量: {result['deleted_files']} 个\n"
            f"删除glass_data记录: {result['deleted_glass_data']} 条\n"
            f"耗时: {execution_time}ms")
        
        # 清空搜索结果
        self.current_results = []
        self.result_text.clear()
        
        self.progress.close()
        self.set_ui_enabled(True)
        self.status_label.setText("删除完成")
    
    def on_delete_error(self, error_msg):
        """删除错误回调"""
        QMessageBox.critical(self, "错误", f"删除时出错: {error_msg}")
        self.progress.close()
        self.set_ui_enabled(True)

# === 主程序入口 ===
if __name__ == "__main__":
    import sys
    
    app = QApplication(sys.argv)
    
    # 设置应用程序信息
    app.setApplicationName("玻璃切割数据管理工具")
    app.setApplicationVersion("2.0.0")
    
    # 设置现代化UI样式
    try:
        app.setStyle("Fusion")  # 使用Fusion主题，更现代化
    except:
        pass
    
    # 设置应用程序样式表
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
    
    try:
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        QMessageBox.critical(None, "启动错误", f"程序启动失败: {e}")
        sys.exit(1)