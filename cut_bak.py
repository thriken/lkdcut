from gettext import find
import os
import sys
import re
import hashlib
import sqlite3
import time
import binascii
from PyQt5.QtGui import QIcon
from datetime import datetime, timedelta
import glob
from xlutils.copy import copy
from xlrd import open_workbook
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QTableWidget,
    QTableWidgetItem, QAbstractItemView,
    QAbstractScrollArea, QToolBar, QAction, QProgressDialog,
    QMessageBox,QButtonGroup,QRadioButton
)
from PyQt5.QtCore import Qt

# ===========================================
# 程序配置 - 方便移植到其他基地
# ===========================================

# 数据库配置
LOCAL_DB_NAME = "glass_data.db"
REMOTE_DB_PATH = r"\\\\landierp\\办公室\\补片程序\\glass_data.db"

# 文件处理配置
WORK_DIRECTORY = r"\\\\landierp\\Share\\激光文件"

# Excel导出配置
EXCEL_TEMPLATE = "origin.xls"
EXPORT_DIRECTORY = os.path.join(os.path.expanduser("~"), "Desktop", "")

# ===========================================
# Utils 功能
def get_current_time():
    """获取当前时间格式 hh:mm:ss"""
    return datetime.now().strftime("%H:%M:%S")

def log_message(message):
    """输出带时间戳的日志信息"""
    timestamp = get_current_time()
    print(f"[{timestamp}] {message}")

# Utils 功能
def get_file_path(root_dir, directory, file_name):
    """统一获取文件路径"""
    if not all([root_dir, file_name]):
        raise ValueError("根目录和文件名不能为空")
    
    # 确保使用正确的网络路径格式
    root_dir = r'\\landierp\Share\激光文件'
    
    if not directory or directory == '.':
        return os.path.normpath(os.path.join(root_dir, file_name))
    
    # 不需要处理路径分隔符，保持原始格式
    if directory:
        full_path = os.path.join(root_dir, directory, file_name)
    else:
        full_path = os.path.join(root_dir, file_name)
        
    return full_path  # 不使用 normpath，保持原始的网络路径格式

def calculate_file_md5(file_path):
    """计算文件MD5值"""
    try:
        file_path = os.path.normpath(file_path)
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return None
            
        try:
            with open(file_path, 'rb') as f:
                md5_hash = hashlib.md5()
                for chunk in iter(lambda: f.read(4096), b""):
                    md5_hash.update(chunk)
                return md5_hash.hexdigest()
        except PermissionError:
            print(f"没有权限访问文件: {file_path}")
            return None
        except Exception as e:
            print(f"读取文件时出错: {file_path}, 错误: {e}")
            return None
    except Exception as e:
        print(f"计算MD5时出错: {e}")
        return None

def calculate_file_crc32(file_path):
    """计算文件CRC32值"""
    try:
        file_path = os.path.normpath(file_path)
        if not os.path.exists(file_path):
            return None
            
        try:
            with open(file_path, 'rb') as f:
                crc = 0
                for chunk in iter(lambda: f.read(4096), b""):
                    crc = binascii.crc32(chunk, crc)
                return format(crc & 0xFFFFFFFF, '08x')
        except Exception:
            return None
    except Exception:
        return None

# 数据库装饰器
def db_connection(func):
    def wrapper(self, *args, **kwargs):
        conn = self.connect()
        cursor = conn.cursor()
        try:
            result = func(self, cursor, *args, **kwargs)
            conn.commit()
            return result
        except Exception as e:
            conn.rollback()
            print(f"数据库操作出错: {e}")
            raise
        finally:
            conn.close()
    return wrapper

# G代码解析器
class GCodeParser:
    def __init__(self):
        self.header_pattern = re.compile(r'N(\d+)\s+P(\d+)\s+=\s+(.+)')
        self.data_pattern = re.compile(r'N(\d+)\s+P(\d+)\s+=\s+(.+)')

    def parse_file(self, file_path):
        try:
            # 优先使用 GB2312 编码
            try:
                with open(file_path, 'r', encoding='gb2312') as f:
                    content = f.readlines()
            except UnicodeDecodeError:
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.readlines()
                except UnicodeDecodeError:
                    log_message(f"无法解析文件编码: {file_path}")
                    return None

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
                        
                        # 在处理客户名称时进行编码转换
                        if processing_pieces and p_num.startswith('4'):
                            current_piece += 1
                            parts = value.split('_')
                            if len(parts) >= 22:
                                piece_data = common_data.copy()
                                # 确保客户名称使用 UTF-8 编码存储
                                customer_name = parts[6]
                                try:
                                    if isinstance(customer_name, str):
                                        customer_name = customer_name.encode('utf-8').decode('utf-8')
                                except UnicodeError:
                                    try:
                                        customer_name = customer_name.encode('gb2312').decode('utf-8')
                                    except UnicodeError:
                                        customer_name = '未知客户'
                                
                                piece_data.update({
                                    'cut_width': int(parts[0]),
                                    'cut_height': int(parts[1]),
                                    'cut_x': int(parts[2]),
                                    'cut_y': int(parts[3]),
                                    'order_width': int(parts[4]),
                                    'order_height': int(parts[5]),
                                    'customer_name': customer_name,  # 使用处理后的客户名称
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

# 数据库管理器
class DatabaseManager:
    def __init__(self):
        self.db_name = "glass_data.db"
        self.remote_db_name = r"\\landierp\办公室\补片程序\glass_data.db"
        # 检查并同步数据库
        self.sync_database()
        self.init_database()

    def sync_database(self):
        """检查并同步本地和远程数据库"""
        try:
            # 检查本地数据库是否存在
            local_exists = os.path.exists(self.db_name)
            remote_exists = os.path.exists(self.remote_db_name)
            
            if not remote_exists:
                log_message("远程数据库不存在，将使用本地数据库")
                return
                
            if not local_exists:
                print("本地数据库不存在，从远程复制")
                self.copy_remote_to_local()
                return
                
            # 获取文件修改时间和大小
            local_mtime = os.path.getmtime(self.db_name)
            remote_mtime = os.path.getmtime(self.remote_db_name)
            
            # 比较修改时间，添加相等的情况判断
            if remote_mtime > local_mtime:
                print("远程数据库较新，正在更新本地数据库...")
                self.copy_remote_to_local()
            elif remote_mtime < local_mtime:
                print(f"本地数据库较新（本地：{datetime.fromtimestamp(local_mtime).strftime('%Y-%m-%d %H:%M:%S')}，远程：{datetime.fromtimestamp(remote_mtime).strftime('%Y-%m-%d %H:%M:%S')}）")
                # 如果本地较新，复制本地数据库到远程
                try:
                    import shutil
                    shutil.copy2(self.db_name, self.remote_db_name)
                    print("已将本地数据库同步到远程")
                except Exception as e:
                    print(f"同步到远程时出错: {e}")
                    print("将继续使用本地数据库")
            else:
                print(f"本地和远程数据库时间相同（{datetime.fromtimestamp(local_mtime).strftime('%Y-%m-%d %H:%M:%S')}），无需同步")
                
        except Exception as e:
            print(f"同步数据库时出错: {e}")
            print("将继续使用本地数据库")

    def copy_remote_to_local(self):
        """从远程复制数据库到本地"""
        try:
            import shutil
            # 如果本地数据库存在，先备份
            if os.path.exists(self.db_name):
                backup_name = f"{self.db_name}.bak"
                shutil.copy2(self.db_name, backup_name)
                print(f"已创建本地数据库备份: {backup_name}")
            
            # 复制远程数据库到本地
            shutil.copy2(self.remote_db_name, self.db_name)
            print("远程数据库已成功复制到本地")
            
        except Exception as e:
            print(f"复制数据库时出错: {e}")
            print("将继续使用本地数据库")

    def connect(self):
        return sqlite3.connect(self.db_name)

    def init_database(self):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS file_info (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            directory TEXT,
            file_name TEXT,
            file_md5 TEXT,
            file_mtime REAL,
            file_size INTEGER,
            file_crc32 TEXT,
            last_processed_time REAL,
            created_time REAL DEFAULT (strftime('%s', 'now')),
            modified_time REAL DEFAULT (strftime('%s', 'now')),
            UNIQUE(directory, file_name)
        )''')

        cursor.execute('''
        CREATE TABLE IF NOT EXISTS glass_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id INTEGER,
            raw_width INTEGER,
            raw_height INTEGER,
            material_code TEXT,
            layout_number INTEGER,
            total_layouts INTEGER,
            thickness INTEGER,
            cut_width INTEGER,
            cut_height INTEGER,
            cut_x INTEGER,
            cut_y INTEGER,
            order_width INTEGER,
            order_height INTEGER,
            customer_name TEXT,
            piece_number INTEGER,
            order_number TEXT,
            dm_code TEXT,
            order_size TEXT,
            reference_edge TEXT,
            group_number TEXT,
            code_3c_position TEXT,
            dm_code_position TEXT,
            tiya_3c_position TEXT,
            FOREIGN KEY(file_id) REFERENCES file_info(id)
        )''')
        
        conn.commit()
        conn.close()

    @db_connection
    def get_file_md5(self, cursor, directory, file_name):
        cursor.execute('''
            SELECT file_md5 
            FROM file_info 
            WHERE directory = ? AND file_name = ?
        ''', (directory, file_name))
        result = cursor.fetchone()
        return result[0] if result else None

    @db_connection
    def get_file_id(self, cursor, directory, file_name):
        cursor.execute('''
            SELECT id FROM file_info 
            WHERE directory = ? AND file_name = ?
        ''', (directory, file_name))
        result = cursor.fetchone()
        return result[0] if result else None

    @db_connection
    def get_file_data(self, cursor, file_id):
        cursor.execute('''
            SELECT * FROM glass_data 
            WHERE file_id = ?
            ORDER BY piece_number
        ''', (file_id,))
        return cursor.fetchall()

    # 新增的搜索方法 按文件名搜索
    @db_connection
    def search_by_filename(self, cursor, filename_pattern):
        """
        通过文件名模式搜索文件及其相关数据
        :param cursor: 数据库游标
        :param filename_pattern: 文件名搜索模式
        :return: 搜索结果列表
        """
        try:
            # 使用 LIKE 进行模糊匹配
            query = '''
                SELECT g.*, f.directory, f.file_name
                FROM glass_data g
                JOIN file_info f ON g.file_id = f.id
                WHERE f.file_name LIKE LOWER(?)
                ORDER BY f.directory, g.layout_number, g.piece_number
            '''
            cursor.execute(query, (f'%{filename_pattern}%',))
            results = cursor.fetchall()
            return results
        except Exception as e:
            print(f"搜索文件时出错: {e}")
            return []
      
    # 新增的搜索方法 按分组号搜索
    @db_connection
    def search_by_group(self, cursor, group_number):
        """
        通过分组号搜索相关数据
        搜索功能增强：采用短横链接2个参数，前面是分组号，后面是文件名
        :param cursor: 数据库游标
        :param group_number: 分组号
        :return: 搜索结果列表
        """
        try:
            group_number = str(group_number) if group_number is not None else ""
            if '-' in group_number:
                # 拆分字符串
                group_number, file_name = group_number.split('-')
                query = '''
                    SELECT g.*, f.directory, f.file_name
                    FROM glass_data g
                    JOIN file_info f ON g.file_id = f.id
                    WHERE LOWER(g.group_number) LIKE LOWER(?) and f.file_name LIKE LOWER(?)
                    ORDER BY f.directory, g.layout_number, g.piece_number
                '''
                cursor.execute(query, (f'%{group_number}%', f'%{file_name}%'))
            else:
                query = '''
                    SELECT g.*, f.directory, f.file_name
                    FROM glass_data g
                    JOIN file_info f ON g.file_id = f.id
                    WHERE LOWER(g.group_number) LIKE LOWER(?)
                    ORDER BY f.directory, g.layout_number, g.piece_number
                '''
                cursor.execute(query, (f'%{group_number}%',))
            results = cursor.fetchall()
            return results
        except Exception as e:
            print(f"搜索分组时出错: {e}")
            return []
    # 新增的搜索方法 按尺寸搜索
    @db_connection
    def get_pieces_by_size(self, cursor, width=None, include_edge=False):
        print(f"\n开始搜索尺寸: {width}, 包含磨边位: {include_edge}")
        
        if width:
            # 确保 width 是字符串类型
            width = str(width)
            # 检查是否包含 'x' 或 'X'
            has_x = 'x' in width.lower()


            if include_edge:  # 勾选磨边位
                if has_x:
                    # 带x的尺寸搜索切割尺寸X和Y
                    dimensions = width.lower().split('x')
                    query = '''
                        SELECT g.*, f.directory, f.file_name
                        FROM glass_data g
                        JOIN file_info f ON g.file_id = f.id
                        WHERE g.cut_width = ? AND g.cut_height = ?
                        ORDER BY f.directory, g.layout_number, g.piece_number
                    '''
                    params = (dimensions[0].strip(), dimensions[1].strip())
                else:
                    # 单数字搜索切割尺寸x或者y
                    query = '''
                        SELECT g.*, f.directory, f.file_name
                        FROM glass_data g
                        JOIN file_info f ON g.file_id = f.id
                        WHERE g.cut_width = ? OR g.cut_height = ?
                        ORDER BY f.directory, g.layout_number, g.piece_number
                    '''
                    params = (width, width)
            else:  # 未勾选磨边位
                if has_x:
                    # 带x的尺寸搜索订单尺寸
                    search_size = f"{width}%"
                    query = '''
                        SELECT g.*, f.directory, f.file_name
                        FROM glass_data g
                        JOIN file_info f ON g.file_id = f.id
                        WHERE g.order_size LIKE ?
                        ORDER BY f.directory, g.layout_number, g.piece_number
                    '''
                    params = (search_size,)
                else:
                    # 单数字搜索订单尺寸
                    search_size = f"%{width}%"
                    query = '''
                        SELECT g.*, f.directory, f.file_name
                        FROM glass_data g
                        JOIN file_info f ON g.file_id = f.id
                        WHERE g.order_size LIKE ?
                        ORDER BY f.directory, g.layout_number, g.piece_number
                    '''
                    params = (search_size,)
        else:
            # 如果没有输入尺寸,返回前300条数据
            query = '''
                SELECT g.*, f.directory, f.file_name
                FROM glass_data g
                JOIN file_info f ON g.file_id = f.id
                ORDER BY f.directory, g.layout_number, g.piece_number
                LIMIT 300
            '''
            params = ()
            
        cursor.execute(query, params)
        results = cursor.fetchall()
        return results

    @db_connection
    def insert_data(self, cursor, directory, file_name, pieces_data):
        """插入数据 - 保持向后兼容"""
        try:
            file_path = get_file_path(os.path.dirname(__file__), directory, file_name)
            file_md5 = calculate_file_md5(file_path)
            
            # 使用时间戳版本插入数据（不需要传递cursor，装饰器会自动处理）
            return self.insert_data_with_timestamp(directory, file_name, pieces_data, file_md5, None, None, None)
        except Exception as e:
            print(f"插入数据时出错: {e}")
            raise

    @db_connection
    def insert_data_with_timestamp(self, cursor, directory, file_name, pieces_data, file_md5, file_mtime, file_size, file_crc32):
        """插入数据并包含完整检查信息"""
        try:
            file_path = get_file_path(os.path.dirname(__file__), directory, file_name)
            
            # 如果未提供时间信息，则从文件获取
            if file_mtime is None:
                try:
                    file_mtime = os.path.getmtime(file_path)
                except:
                    file_mtime = time.time()
            
            if file_size is None:
                try:
                    file_size = os.path.getsize(file_path)
                except:
                    file_size = 0
            
            if file_crc32 is None:
                file_crc32 = calculate_file_crc32(file_path)
            
            # 先删除该文件的glass_data记录
            cursor.execute('''
                DELETE FROM glass_data 
                WHERE file_id IN (
                    SELECT id FROM file_info 
                    WHERE directory = ? AND file_name = ?
                )
            ''', (directory, file_name))
            
            # 删除文件信息
            cursor.execute('''
                DELETE FROM file_info 
                WHERE directory = ? AND file_name = ?
            ''', (directory, file_name))
            
            # 插入文件信息（包含完整检查信息）
            cursor.execute('''
                INSERT INTO file_info (directory, file_name, file_md5, file_mtime, file_size, file_crc32, last_processed_time)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (directory, file_name, file_md5, file_mtime, file_size, file_crc32, time.time()))
            file_id = cursor.lastrowid
            
            # 插入玻璃数据
            for piece in pieces_data:
                cursor.execute('''
                    INSERT INTO glass_data (
                        file_id, raw_width, raw_height, material_code,
                        layout_number, total_layouts, thickness,
                        cut_width, cut_height, cut_x, cut_y,
                        order_width, order_height, customer_name,
                        piece_number, order_number, dm_code,
                        order_size, reference_edge, group_number,
                        code_3c_position, dm_code_position, tiya_3c_position
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    file_id,
                    piece.get('raw_width'),
                    piece.get('raw_height'),
                    piece.get('material_code'),
                    piece.get('layout_number'),
                    piece.get('total_layouts'),
                    piece.get('thickness'),
                    piece.get('cut_width'),
                    piece.get('cut_height'),
                    piece.get('cut_x'),
                    piece.get('cut_y'),
                    piece.get('order_width'),
                    piece.get('order_height'),
                    piece.get('customer_name'),
                    piece.get('piece_number'),
                    piece.get('order_number'),
                    piece.get('dm_code'),
                    piece.get('order_size'),
                    piece.get('reference_edge'),
                    piece.get('group_number'),
                    piece.get('code_3c_position'),
                    piece.get('dm_code_position'),
                    piece.get('tiya_3c_position')
                ))
            return True
        except Exception as e:
            print(f"插入数据时出错: {e}")
            raise

# Excel导出器
class ExcelExporter:
    def __init__(self):
        self.template_path = os.path.join(os.path.dirname(__file__), 'origin.xls')
        # 导出目录 设置为桌面
        self.export_dir  = os.path.join(os.path.expanduser("~"), "Desktop", "")
        # 导出目录 设置为当前目录
        #self.export_dir = os.path.dirname(__file__)
        self.template_columns = {
            'A': '材料代码',
            'B': '切片尺寸X',
            'C': '切片尺寸Y',
            'D': '切片数量',
            'E': '订单号',
            'E': '客户名',
            'R': 'DM码内容',
            'U': '订单尺寸',
            'V': '基准边',
            'W': '分组号',
            'X': '3C码号和位置码',
            'Y': 'DM码号和位置码',
            'Z': '蒂亚科技专用3C码号和位置码'
        }
        self.last_export_file = None
        self.last_material_code = None  # 添加新属性来存储上次导出的原始材料代码

    def get_export_filename(self, material_code, total_pieces=None):
        # 保存原始材料代码
        self.last_material_code = material_code
        today = datetime.now().strftime('%m%d')
        if total_pieces:
            return f"{material_code}-{total_pieces}-{today}.xls"
        else:
            pattern = f"{material_code}-{today}-*.xls"
            existing_files = glob.glob(os.path.join(self.export_dir, pattern))
            if not existing_files:
                return f"{material_code}-{today}-1.xls"
            
            max_num = max([int(f.split('-')[-1].split('.')[0]) for f in existing_files])
            return f"{material_code}-{today}-{max_num + 1}.xls"

    def export_data(self, data, material_code, total_pieces=None, append=False):
        print(f"开始导出数据: material_code={material_code}, total_pieces={total_pieces}, append={append}")
        try:
            if append and self.last_export_file and os.path.exists(self.last_export_file):
                print(f"检测到上次导出文件: {self.last_export_file}")
                print(f"上次材料代码: {self.last_material_code}, 当前材料代码: {material_code}")
                
                # 检查材料代码是否匹配
                if self.last_material_code and material_code != self.last_material_code:
                    QMessageBox.warning(
                        None,
                        '材料不匹配',
                        f'当前导出的材料({material_code})与上一个导出文件的材料({self.last_material_code})不同，无法追加导出。'
                    )
                    print(f"材料不匹配：当前材料 {material_code}，上次材料 {self.last_material_code}")
                    self.material_mismatch = True  # 设置状态标志
                    return False  # 材料不匹配时终止导出操作，但不再显示额外的导出失败提示
                
                # 直接使用上次的导出文件
                output_path = self.last_export_file
                rb = open_workbook(output_path, formatting_info=True)

            else:
                print("创建新导出文件")
                output_path = os.path.join(
                    self.export_dir,
                    self.get_export_filename(material_code, total_pieces)
                )
                print(f"导出文件路径: {output_path}")
                try:
                    rb = open_workbook(self.template_path, formatting_info=True)
                except Exception as e:
                    print(f"读取模板文件失败: {e}")
                    return False

            wb = copy(rb)
            ws = wb.get_sheet(0)
            
            sheet = rb.sheet_by_index(0)
            
            if append and sheet.nrows > 1:
                start_row = 1
                last_row = 1
                for row in range(1, sheet.nrows):
                    if sheet.cell(row, 0).value:
                        last_row = row
                start_row = last_row + 1
                print(f"找到最后一条数据在第 {last_row} 行")
            else:
                start_row = 1
                
            print(f"数据将写入第 {start_row} 行")
            
            for index, row in enumerate(data):
                try:
                    current_row = start_row + index
                    print(f"写入第 {current_row} 行数据")
                    # 材料代码由正则表达式提取成 mm格式
                    thickness = ''
                    if material_code:
                        # 使用正则表达式提取开头的数字
                        match = re.match(r'(\d+)', material_code)
                        if match:
                            thickness = f"{match.group(1)}mm"
                    ws.write(current_row, 0, thickness)
                    #ws.write(current_row, 0, row.get('material_code', ''))
                    ws.write(current_row, 1, row.get('cut_width', ''))
                    ws.write(current_row, 2, row.get('cut_height', ''))
                    ws.write(current_row, 3, 1)
                    ws.write(current_row, 4, row.get('order_number', ''))
                    ws.write(current_row, 5, row.get('customer_name', ''))
                    ws.write(current_row, 17, row.get('dm_code', ''))
                    ws.write(current_row, 20, row.get('order_size', ''))
                    ws.write(current_row, 21, row.get('reference_edge', ''))
                    ws.write(current_row, 22, row.get('group_number', ''))
                    ws.write(current_row, 23, row.get('code_3c_position', ''))
                    ws.write(current_row, 24, row.get('dm_code_position', ''))
                    ws.write(current_row, 25, row.get('tiya_3c_position', ''))
                    print(f"已处理第 {index + 1}/{len(data)} 行数据")
                except Exception as e:
                    print(f"处理第 {index + 1} 行数据时出错: {e}")
                    continue

            print("保存Excel文件")
            wb.save(output_path)
            self.last_material_code = material_code
            self.last_export_file = output_path
            print(f"文件已保存: {output_path}")
            return True
            
        except Exception as e:
            print(f"导出Excel时发生错误: {e}")
            return False

# 主窗口
# 程序功能：
# 1. 扫描文件夹
# 2. 解析G代码
# 3. 搜索数据尺寸 文件名 分组号
# 4. 导出Excel
# 5. 数据库管理SQLITE3

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.parser = GCodeParser()
        self.db = DatabaseManager()
        self.exporter = ExcelExporter()
        self.default_dir = r'\\landierp\Share\激光文件'
        
        # 添加状态栏
        self.statusBar()
        self.scan_status = QLabel()
        self.search_status = QLabel()
        # 新增数据库更新时间标签
        self.db_update_status = QLabel()  
         # 设置状态栏样式
        status_style = """
            QLabel {
                color: #ff4d4f;
                font-size: 14px;
                font-weight: bold;
                padding: 5px;
                font-family: "Microsoft YaHei UI";
            }
        """
        self.scan_status.setStyleSheet(status_style)
        self.search_status.setStyleSheet(status_style)
        self.db_update_status.setStyleSheet(status_style)

        self.statusBar().addWidget(self.scan_status)
        self.statusBar().addWidget(self.search_status)
        # 使用addPermanentWidget将标签添加到最右侧
        self.statusBar().addPermanentWidget(self.db_update_status)  
        # 更新数据库修改时间显示
        self.update_db_status()

        # 创建表格
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            '目录', '文件名', '客户名称', 'DM码','订单尺寸',
             '分组号', '3C码位置', 'DM码位置'
        ])
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # 设置表格为只读模式
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # 设置表格样式
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                alternate-background-color: #f5f5f5;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                gridline-color: #e8e8e8;
            }
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f0f0f0;
            }
            QTableWidget::item:selected {
                background-color: #e6f7ff;
                color: black;
            }
            QHeaderView::section {
                background-color: #fafafa;
                padding: 8px;
                border: none;
                border-bottom: 2px solid #e8e8e8;
                font-weight: bold;
                color: #262626;
            }
            QHeaderView::section:hover {
                background-color: #f0f0f0;
            }
        """)
        
        # 设置列宽
        self.table.setColumnWidth(0, 220)  # 目录
        self.table.setColumnWidth(1, 220)  # 文件名
        self.table.setColumnWidth(2, 120)  # 客户名称
        self.table.setColumnWidth(3, 100)  # DM码
        self.table.setColumnWidth(4, 100)  # 订单尺寸
        self.table.setColumnWidth(5, 80)   # 基准边
        self.table.setColumnWidth(6, 100)  # 分组号
        self.table.setColumnWidth(7, 100)  # 3C码位置
        self.table.setColumnWidth(8, 80)   # DM码位置
        
        # 启用隔行变色
        self.table.setAlternatingRowColors(True)
        
        # 设置表头属性
        header = self.table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header.setStretchLastSection(True)  # 最后一列自动填充
        
        # 设置垂直表头隐藏
        self.table.verticalHeader().setVisible(False)
        
        # 设置表格自动调整大小策略
        self.table.setSizeAdjustPolicy(QAbstractScrollArea.AdjustToContents)
        
        # 设置表格的选择模式为整行选择
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
       
        self.init_ui()

    def update_db_status(self):
        """更新数据库修改时间显示"""
        try:
            #db_path = os.path.join(os.path.dirname(__file__), "glass_data.db")
            db_path =  "glass_data.db"
            if os.path.exists(db_path):
                mtime = os.path.getmtime(db_path)
                mtime_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
                self.db_update_status.setText(f"数据更新时间：{mtime_str}")
            else:
                self.db_update_status.setText("数据库文件不存在")
        except Exception as e:
            self.db_update_status.setText("无法获取更新时间")

    def init_ui(self):
        # 设置窗口基本属性
        self.setWindowTitle('玻璃切割数据管理')
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowIcon(QIcon("./icon.ico"))
        # 创建主布局
        main_layout = QVBoxLayout()
        main_layout.setSpacing(10)  # 设置布局间距
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # 创建顶部工具栏并固定
        toolbar = QToolBar()
        toolbar.setMovable(False)  # 禁止移动
        toolbar.setFloatable(False)  # 禁止浮动
        self.addToolBar(toolbar)
        toolbar.setStyleSheet("""
            QToolBar {
                background: #ffffff;
                border-bottom: 1px solid #e0e0e0;
                padding: 5px;
                font-family: "Microsoft YaHei UI";
            }
            QToolButton {
                padding: 5px;
                border: none;
                border-radius: 4px;
            }
            QToolButton:hover {
                background: #e6f7ff;
            }
        """)

        # 添加扫描按钮
        scan_action = QAction('刷新数据', self)
        scan_action.triggered.connect(self.scan_files)
        toolbar.addAction(scan_action)

        # 创建搜索区域（固定高度的widget）
        search_widget = QWidget()
        search_widget.setFixedHeight(50)  # 固定高度
        search_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                font-family: "Microsoft YaHei UI";
            }
            QPushButton {
                min-width: 90px;
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
            
            QCheckBox {
                margin: 0 15px;
                color: #595959;
                font-weight: 500;
                font-size: 14px;
                
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }

            QLabel {
                color: #262626;
                font-weight: 500;
            }
        """)
        
        search_layout = QHBoxLayout(search_widget)
        search_layout.setContentsMargins(10, 5, 10, 5)  # 设置边距

        # 创建搜索类型单选按钮组
        self.search_type_group = QButtonGroup(self)
        
        # 创建单选按钮
        self.size_radio = QRadioButton("尺寸")
        self.filename_radio = QRadioButton("文件名")
        self.group_radio = QRadioButton("架号")
        self.size_radio.setContentsMargins(0, 0, 10, 0)  # 只设置右边距10px
        self.filename_radio.setContentsMargins(0, 0, 10, 0)
        self.group_radio.setContentsMargins(0, 0, 10, 0)
        # 设置默认选中尺寸搜索
        self.size_radio.setChecked(True)
        
        # 将单选按钮添加到按钮组
        self.search_type_group.addButton(self.size_radio)
        self.search_type_group.addButton(self.filename_radio)
        self.search_type_group.addButton(self.group_radio)
        
        # 将单选按钮添加到搜索布局
        search_layout.addWidget(self.size_radio)
        search_layout.addWidget(self.filename_radio)
        search_layout.addWidget(self.group_radio)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('请输入搜索内容')
        self.edge_checkbox = QCheckBox('含磨边位')
        search_button = QPushButton('搜索')
        search_button.clicked.connect(self.search_data)
        
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.edge_checkbox)
        search_layout.addWidget(search_button)

        self.search_input.returnPressed.connect(self.search_data)

        search_layout.addStretch()
        
        # 导出按钮
        export_button = QPushButton('导出选中')
        export_button.clicked.connect(self.export_excel)
        search_layout.addWidget(export_button)
        
        # 追加导出按钮
        append_export_button = QPushButton('追加导出')
        append_export_button.clicked.connect(lambda: self.export_excel(True))
        search_layout.addWidget(append_export_button)
        
        main_layout.addWidget(search_widget)

        # 添加主程序入口
        # 添加表格到主布局
        main_layout.addWidget(self.table)

    def scan_files(self):
        # 创建进度对话框
        progress = QProgressDialog("准备扫描文件...", "取消", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setAutoClose(True)
        progress.setAutoReset(True)
        
        # 更新进度对话框的样式
        progress.setStyleSheet("""
            QProgressDialog {
                background-color: #ffffff;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                min-width: 400px;
                font-family: "Microsoft YaHei UI";
            }
            QProgressBar {
                border: 1px solid #d9d9d9;
                border-radius: 3px;
                text-align: center;
                background-color: #f5f5f5;
            }
            QProgressBar::chunk {
                background-color: #1890ff;
                border-radius: 2px;
            }
        """)
        
        # 显示进度对话框
        progress.show()
        QApplication.processEvents()
        
        # 调用处理目录的方法
        self.process_directory(self.default_dir, progress)
        self.load_data()
        self.update_db_status()


    def process_directory(self, workdir, progress):
        print(f"开始处理目录: {workdir}")
        total_files = 0
        new_files = 0
        updated_files = 0
        unchanged_files = 0
        total_pieces = 0
        changed_pieces = 0

        # 性能计时变量
        phase_times = {
            'file_walk': 0.0,
            'md5_calculation': 0.0,
            'database_queries': 0.0,
            'gcode_parsing': 0.0,
            'database_insertion': 0.0,
            'total': 0.0
        }
        
        phase_times['total'] = time.time()

        try:
            # 阶段1: 文件遍历和统计
            phase_start = time.time()
            progress.setLabelText("正在统计文件数量...")
            QApplication.processEvents()

            # 获取数据库中所有文件信息（包括时间戳）
            conn = self.db.connect()
            cursor = conn.cursor()
            cursor.execute('SELECT directory, file_name, file_md5, last_processed_time FROM file_info')
            existing_files = {(row[0], row[1]): (row[2], row[3]) for row in cursor.fetchall()}
            conn.close()
            
            # 先统计总文件数并收集需要处理的文件
            g_files = []
            for root, _, files in os.walk(workdir):
                for file in files:
                    if file.endswith('.g'):
                        relative_dir = os.path.relpath(root, workdir)
                        if relative_dir == '.':
                            relative_dir = ''
                        
                        file_path = os.path.join(root, file)
                        g_files.append((file_path, file, relative_dir))
                        total_files += 1
            
            phase_times['file_walk'] = time.time() - phase_start
            print(f"阶段1 - 文件遍历完成: {phase_times['file_walk']:.3f}秒")

            # 检查文件是否已存在
            if not os.path.exists(workdir):
                print(f"无法访问目录: {workdir}")
                progress.close() 
                return

            # 更新进度对话框
            total_process_files = len(g_files)
            if total_process_files == 0:
                progress.setValue(100)
                progress.close()
                status_text = f"扫描完成 - 总文件: {total_files} | 新增: {new_files} | 更新: {updated_files} | 未变化: {unchanged_files}"
                self.scan_status.setText(status_text)
                return

            # 阶段2: 处理文件
            progress.setValue(0)
            progress.setLabelText("正在处理文件...")
            
            # 优化：先使用文件修改时间作为初步筛选，避免不必要的MD5计算
            need_process_files = []
            db_check_start = time.time()
            
            # 获取数据库中所有文件的修改时间（如果存在）
            file_time_cache = {}
            conn = self.db.connect()
            cursor = conn.cursor()
            cursor.execute("SELECT directory, file_name, file_md5, last_processed_time FROM file_info")
            for row in cursor.fetchall():
                file_time_cache[(row[0], row[1])] = (row[2], row[3])
            conn.close()
            
            phase_times['database_queries'] += time.time() - db_check_start
            
            # 使用智能策略进行快速筛选
            time_filter_start = time.time()
            for file_path, file, directory in g_files:
                existing_info = file_time_cache.get((directory, file))
                
                if not existing_info:
                    # 新文件需要处理
                    need_process_files.append((file_path, file, directory, None))
                else:
                    existing_md5, last_processed_time = existing_info
                    
                    # 检查文件是否在一周内处理过（智能文件检查策略）
                    one_week_ago = time.time() - (7 * 24 * 3600)
                    
                    if last_processed_time and last_processed_time > one_week_ago:
                        # 一周内处理过的文件，可以跳过MD5计算（除非文件修改时间变化）
                        try:
                            file_mtime = os.path.getmtime(file_path)
                            # 如果文件修改时间变化，仍然需要检查
                            need_process_files.append((file_path, file, directory, existing_md5))
                        except:
                            # 如果无法获取文件时间，保守处理
                            need_process_files.append((file_path, file, directory, existing_md5))
                    else:
                        # 超过一周未处理的文件，需要检查
                        need_process_files.append((file_path, file, directory, existing_md5))
            
            phase_times['database_queries'] += time.time() - time_filter_start
            
            # 处理需要更新的文件
            for current_count, (file_path, file, directory, existing_md5) in enumerate(need_process_files, 1):
                if progress.wasCanceled():
                    break
                    
                progress_value = int((current_count / len(need_process_files)) * 100)
                progress.setValue(progress_value)
                progress.setLabelText(f"正在处理文件 ({current_count}/{len(need_process_files)}): {file}")
                
                try:
                    # 优化：只有在数据库中存在记录时才计算MD5进行对比
                    if existing_md5 is not None:
                        # 阶段2.1: MD5计算（仅对已存在文件）
                        md5_start = time.time()
                        current_md5 = calculate_file_md5(file_path)
                        phase_times['md5_calculation'] += time.time() - md5_start
                        
                        if not current_md5:
                            print(f"MD5计算失败，跳过文件: {file}")
                            continue
                        
                        # 检查文件是否需要处理
                        if existing_md5 == current_md5:
                            unchanged_files += 1
                            continue
                    else:
                        # 新文件，需要计算MD5
                        md5_start = time.time()
                        current_md5 = calculate_file_md5(file_path)
                        phase_times['md5_calculation'] += time.time() - md5_start
                        
                        if not current_md5:
                            print(f"MD5计算失败，跳过文件: {file}")
                            continue
                    
                    # 阶段2.3: G代码解析
                    parsing_start = time.time()
                    pieces_data = self.parser.parse_file(file_path)
                    phase_times['gcode_parsing'] += time.time() - parsing_start
                    
                    if pieces_data:
                        current_pieces = len(pieces_data)
                        total_pieces += current_pieces
                        
                        # 阶段2.4: 数据库插入
                        db_insert_start = time.time()
                        if self.db.insert_data(directory, file, pieces_data):
                            if existing_md5 is None:
                                new_files += 1
                                changed_pieces += current_pieces
                            else:
                                updated_files += 1
                                old_pieces = len(self.db.get_file_data(self.db.get_file_id(directory, file)))
                                changed_pieces += abs(current_pieces - old_pieces)
                        phase_times['database_insertion'] += time.time() - db_insert_start
                        
                        print(f">>> 文件 {file} 包含 {current_pieces} 个小片")
                        print(f">>> 当前累计小片数量: {total_pieces}")
                        print("-" * 30)
                    else:
                        print(f"文件解析失败，跳过: {file}")
                        
                except Exception as e:
                    print(f"处理文件时出错: {file}, 错误: {e}")
                    continue
                
                if current_count % 5 == 0:  # 每处理5个文件刷新一次
                    QApplication.processEvents()
            
            phase_times['total'] = time.time() - phase_times['total']
            
            # 输出性能分析结果
            print("\n" + "="*60)
            print("性能分析报告:")
            print("="*60)
            print(f"总耗时: {phase_times['total']:.3f}秒")
            print(f"文件遍历: {phase_times['file_walk']:.3f}秒 ({phase_times['file_walk']/phase_times['total']*100:.1f}%)")
            print(f"MD5计算: {phase_times['md5_calculation']:.3f}秒 ({phase_times['md5_calculation']/phase_times['total']*100:.1f}%)")
            print(f"数据库查询: {phase_times['database_queries']:.3f}秒 ({phase_times['database_queries']/phase_times['total']*100:.1f}%)")
            print(f"G代码解析: {phase_times['gcode_parsing']:.3f}秒 ({phase_times['gcode_parsing']/phase_times['total']*100:.1f}%)")
            print(f"数据库插入: {phase_times['database_insertion']:.3f}秒 ({phase_times['database_insertion']/phase_times['total']*100:.1f}%)")
            print("="*60)
            
            # 识别性能瓶颈
            max_phase = max(phase_times, key=phase_times.get)
            print(f"性能瓶颈: {max_phase} ({phase_times[max_phase]:.3f}秒)")
            
            if max_phase == 'md5_calculation':
                print("优化建议: 考虑使用更快的哈希算法或减少MD5计算的频率")
            elif max_phase == 'gcode_parsing':
                print("优化建议: 优化G代码解析算法，减少正则表达式匹配次数")
            elif max_phase == 'database_insertion':
                print("优化建议: 使用批量插入代替逐条插入，或优化数据库索引")
            elif max_phase == 'database_queries':
                print("优化建议: 减少数据库查询次数，使用缓存机制")
                
            print("="*60)
            
            print("\n处理完成!")
            print(f"总文件数: {total_files}")
            print(f"新文件数: {new_files}")
            print(f"更新文件数: {updated_files}")
            print(f"未变化文件数: {unchanged_files}")
            print(f"总小片数: {total_pieces}")
            print(f"变化小片数: {changed_pieces}")
            
            # 更新状态栏显示扫描结果
            status_text = f"扫描完成 - 总文件: {total_files} | 新增: {new_files} | 更新: {updated_files} | 未变化: {unchanged_files} | 总小片: {total_pieces} | 变化小片: {changed_pieces}"
            self.scan_status.setText(status_text)
            
        except Exception as e:
            error_msg = f"处理目录时出错: {e}"
            print(error_msg)
            self.scan_status.setText("扫描出错")

    def focus_y_input(self):
        self.search_input_y.setFocus()

    def search_data(self):
        try:
            search_text = self.search_input.text().strip()
            
            # 根据选中的单选按钮执行不同的搜索
            if self.size_radio.isChecked():
                # 尺寸搜索
                if not search_text:
                    results = self.db.get_pieces_by_size(None, self.edge_checkbox.isChecked())
                else:
                    if 'x' in search_text.lower():
                        try:
                            width, height = map(int, search_text.lower().split('x'))
                            size = f"{width}x{height}"
                        except ValueError:
                            QMessageBox.warning(self, '警告', '请输入正确的尺寸格式，例如: 1000x2000')
                            return
                    else:
                        try:
                            size = int(search_text)
                        except ValueError:
                            QMessageBox.warning(self, '警告', '请输入正确的尺寸格式')
                            return
                    results = self.db.get_pieces_by_size(size, self.edge_checkbox.isChecked())
            elif self.filename_radio.isChecked():
                # 文件名搜索
                results = self.db.search_by_filename(search_text)
            else:
                # 架号搜索
                results = self.db.search_by_group(search_text)
            self.table.setRowCount(len(results))
            for row, data in enumerate(results):
                for col in range(8):
                    item = QTableWidgetItem()
                    item.setTextAlignment(Qt.AlignCenter)
                    if col == 0:
                        item.setText(str(data[24]))  # directory
                    elif col == 1:
                        item.setText(str(data[25]))  # file_name
                    elif col == 2:
                        item.setText(str(data[14]))  # customer_name
                    elif col == 3:
                        item.setText(str(data[17]))  # DM_code
                    elif col == 4:
                        item.setText(str(data[18]))  # order_size
                    elif col == 5:
                        item.setText(str(data[20]))  # reference_edge
                    elif col == 6:
                        item.setText(str(data[21]))  # code_3c_position
                    elif col == 7:
                        item.setText(str(data[22]))  # dm_code_position
                    self.table.setItem(row, col, item)
            
            # 更新状态栏显示搜索结果
            self.search_status.setText(f"搜索完成 - 找到 {len(results)} 条记录")
            
        except ValueError:
            QMessageBox.warning(self, '警告', '请输入有效的数字')
        except Exception as e:
            QMessageBox.critical(self, '错误', f'搜索时出错: {str(e)}')

    def export_excel(self, append=False):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, '警告', '请先选择要导出的数据')
            return
            
        try:
            export_data = []
            for row in selected_rows:
                row_index = row.row()
                directory = self.table.item(row_index, 0).text()
                file_name = self.table.item(row_index, 1).text()
                customer_name = self.table.item(row_index, 2).text()
                dm_code = self.table.item(row_index, 3).text()
                order_size = self.table.item(row_index, 4).text()
                reference_edge = self.table.item(row_index, 5).text()
                code_3c_position = self.table.item(row_index, 6).text()
                dm_code_position = self.table.item(row_index, 7).text()
                
                # 修改尺寸解析逻辑
                try:
                    # 处理可能的尺寸格式，如 "1075/A" 或 "1075x1194"
                    size_parts = order_size.split('x' if 'x' in order_size else '/')
                    cut_width = int(size_parts[0].strip())
                    cut_height = cut_width  # 如果是正方形，高度等于宽度
                    if len(size_parts) > 1 and size_parts[1].strip().isdigit():
                        cut_height = int(size_parts[1].strip())
                except (ValueError, IndexError):
                    cut_width = 0
                    cut_height = 0
                
                piece_dict = {
                    'material_code': '',
                    'cut_width': 0,  # 初始化为0，稍后从数据库获取实际切片尺寸
                    'cut_height': 0,  # 初始化为0，稍后从数据库获取实际切片尺寸
                    'order_number': '',
                    'dm_code': dm_code,
                    'order_size': order_size,
                    'reference_edge': '',
                    'group_number': '',
                    'code_3c_position': code_3c_position,
                    'dm_code_position': dm_code_position,
                    'tiya_3c_position': ''
                }
                
                # 从数据库获取额外信息
                file_id = self.db.get_file_id(directory, file_name)
                if file_id:
                    file_data = self.db.get_file_data(file_id)
                    if file_data:
                        # 只获取匹配当前选中行的数据
                        for piece in file_data:
                            if (piece[17] == dm_code and  # dm_code
                                piece[18] == order_size):  # order_size
                                piece_dict.update({
                                    'material_code': piece[4],
                                    'cut_width': piece[8],   # 从数据库获取实际切片X尺寸
                                    'cut_height': piece[9],  # 从数据库获取实际切片Y尺寸
                                    'customer_name': piece[14],
                                    'order_number': piece[16],
                                    'reference_edge': piece[19],
                                    'group_number': piece[20],
                                    'tiya_3c_position': piece[23]
                                })
                                break
                
                export_data.append(piece_dict)
            
            if export_data:
                # 获取第一条数据的材料代码
                material_code = export_data[0].get('material_code', '')
                result = self.exporter.export_data(export_data, material_code, len(export_data), append)
                if result:
                    QMessageBox.information(self, '成功', '数据导出成功')
                elif not self.exporter.material_mismatch:  # 只有在非材料不匹配的情况下才显示导出失败
                    QMessageBox.warning(self, '警告', '数据导出失败')
                self.exporter.material_mismatch = False  # 重置状态标志
            else:
                QMessageBox.warning(self, '警告', '没有找到可导出的数据')
                
        except Exception as e:
            QMessageBox.critical(self, '错误', f'导出时出错: {str(e)}')

    def load_data(self):
        # 从数据库加载并显示数据
        pass
            # 添加主程序入口
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
