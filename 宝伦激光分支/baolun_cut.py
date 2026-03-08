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
LOCAL_DB_NAME = "baolun_glass_data.db"
REMOTE_DB_PATH = r"\\hhjy-002\e\切割任务文件夹\baolun_glass_data.db"

# 文件处理配置
WORK_DIRECTORY = r"\\hhjy-002\e\切割任务文件夹"

print(f"WORK_DIRECTORY: {WORK_DIRECTORY}")
print(f"REMOTE_DB_PATH: {REMOTE_DB_PATH}")
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
def get_file_path(directory, file_name):
    """统一获取文件路径"""
    if not file_name:
        raise ValueError("文件名不能为空")
    
    if not directory or directory == '.':
        return os.path.normpath(os.path.join(WORK_DIRECTORY, file_name))
    
    # 不需要处理路径分隔符，保持原始格式
    full_path = os.path.join(WORK_DIRECTORY, directory, file_name)
    return full_path  # 不使用 normpath，保持原始的网络路径格式

def calculate_file_md5(file_path):
    """计算文件MD5值"""
    try:
        file_path = os.path.normpath(file_path)
        
        if not os.path.exists(file_path):
            return None
        
        try:
            with open(file_path, 'rb') as f:
                md5_hash = hashlib.md5()
                for chunk in iter(lambda: f.read(4096), b""):
                    md5_hash.update(chunk)
                return md5_hash.hexdigest()
        except PermissionError:
            print(f"文件权限错误: {file_path}")
            return None
        except FileNotFoundError:
            print(f"文件未找到: {file_path}")
            return None
        except Exception:
            return None
    except Exception:
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
class BaolunGCodeParser:
    def __init__(self):
        # 支持多种格式：有或没有空格的=号
        self.header_pattern = re.compile(r'N(\d+)\s+P(\d+)\s*=\s*(.+)')
        self.data_pattern = re.compile(r'N(\d+)\s+P(\d+)\s*=\s*(.+)')

    def parse_file(self, file_path):
        try:
            # 优先使用 GB2312 (ANSI) 编码
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
            except Exception:
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
                        # N13-N16 行都可能是小片数据，根据P编号判断
                        if n_num in ['13', '14', '15', '16'] and p_num.startswith('4'):
                            processing_pieces = True
                        
                        # 处理小片数据
                        if processing_pieces and p_num.startswith('4'):
                            current_piece += 1
                            parts = value.split('_')
                            
                            # 宝伦格式解析 - 至少28个字段
                            if len(parts) >= 28:
                                piece_data = common_data.copy()
                                
                                # 宝伦1-28字段映射
                                piece_data.update({
                                    'bl1_size_x': float(parts[0]) if parts[0] else 0,      # 宝伦1: 小片X方向尺寸
                                    'bl2_size_y': float(parts[1]) if parts[1] else 0,      # 宝伦2: 小片Y方向尺寸
                                    'bl3_x_coordinate': float(parts[2]) if parts[2] else 0,  # 宝伦3: 左下角X坐标
                                    'bl4_y_coordinate': float(parts[3]) if parts[3] else 0,  # 宝伦4: 左下角Y坐标
                                    'bl5_display_length': float(parts[4]) if parts[4] else 0,  # 宝伦5: 显示的长(X)
                                    'bl6_display_width': float(parts[5]) if parts[5] else 0,   # 宝伦6: 显示的宽(Y)
                                    'bl7_customer': parts[6] if len(parts) > 6 else '',         # 宝伦7: 客户名
                                    'bl8_position': int(parts[7]) if len(parts) > 7 and parts[7].isdigit() else 0,  # 宝伦8: 定位号
                                    'bl9_group': parts[8] if len(parts) > 8 else '',          # 宝伦9: 货架号分组号
                                    'bl10_order': parts[9] if len(parts) > 9 else '',         # 宝伦10: 订单号
                                    'bl11_order_size': parts[10] if len(parts) > 10 else '',   # 宝伦11: 订单尺寸加片标记
                                    'bl12': parts[11] if len(parts) > 11 else '',             # 宝伦12
                                    'bl13': parts[12] if len(parts) > 12 else '',             # 宝伦13
                                    'bl14': parts[13] if len(parts) > 13 else '',             # 宝伦14
                                    'bl15_barcode': parts[14] if len(parts) > 14 else '',      # 宝伦15: 条码号
                                    'bl16_count': int(parts[15]) if len(parts) > 15 and parts[15].isdigit() else 1,  # 宝伦16: 小片数量
                                    'bl17_template1': int(parts[16]) if len(parts) > 16 and parts[16].isdigit() else 0,  # 宝伦17: 标签1激光模板序号
                                    'bl18_edge1': float(parts[17]) if len(parts) > 17 and parts[17].replace('.','').isdigit() else 0,  # 宝伦18: 标签1基准边
                                    'bl19_corner1': int(parts[18]) if len(parts) > 18 and parts[18].isdigit() else 0,  # 宝伦19: 标签1角位号
                                    'bl20_position1': int(parts[19]) if len(parts) > 19 and parts[19].isdigit() else 0,  # 宝伦20: 标签1离边角位置
                                    'bl21_template2': int(parts[20]) if len(parts) > 20 and parts[20].isdigit() else 0,  # 宝伦21: 标签2激光模板序号
                                    'bl22_edge2': float(parts[21]) if len(parts) > 21 and parts[21].replace('.','').isdigit() else 0,  # 宝伦22: 标签2基准边
                                    'bl23_corner2': int(parts[22]) if len(parts) > 22 and parts[22].isdigit() else 0,  # 宝伦23: 标签2角位号
                                    'bl24_position2': int(parts[23]) if len(parts) > 23 and parts[23].isdigit() else 0,  # 宝伦24: 标签2离边角位置
                                    'bl25_template3': int(parts[24]) if len(parts) > 24 and parts[24].isdigit() else 0,  # 宝伦25: 标签3激光模板序号
                                    'bl26_edge3': float(parts[25]) if len(parts) > 25 and parts[25].replace('.','').isdigit() else 0,  # 宝伦26: 标签3基准边
                                    'bl27_corner3': int(parts[26]) if len(parts) > 26 and parts[26].isdigit() else 0,  # 宝伦27: 标签3角位号
                                    'bl28_position3': int(parts[27]) if len(parts) > 27 and parts[27].isdigit() else 0,  # 宝伦28: 标签3离边角位置
                                })
                                pieces_data.append(piece_data)
                        elif not processing_pieces:
                            # 处理通用数据
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
                    except (ValueError, IndexError) as e:
                        print(f"解析数据时出错: {value}, 错误: {e}")
                        continue
                elif line.startswith('N13') and not line.startswith('G'):
                    continue
            return pieces_data

        except Exception as e:
            print(f"解析文件时出错: {e}")
            return None

# 数据库管理器
class DatabaseManager:
    def __init__(self):
        self.db_name = LOCAL_DB_NAME
        self.remote_db_name = REMOTE_DB_PATH
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
                
            if not local_exists:
                print("本地数据库不存在，从远程复制")
                self.copy_remote_to_local()
                
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
            bl1_size_x REAL,
            bl2_size_y REAL,
            bl3_x_coordinate REAL,
            bl4_y_coordinate REAL,
            bl5_display_length REAL,
            bl6_display_width REAL,
            bl7_customer TEXT,
            bl8_position INTEGER,
            bl9_group TEXT,
            bl10_order TEXT,
            bl11_order_size TEXT,
            bl12 TEXT,
            bl13 TEXT,
            bl14 TEXT,
            bl15_barcode TEXT,
            bl16_count INTEGER,
            bl17_template1 INTEGER,
            bl18_edge1 REAL,
            bl19_corner1 INTEGER,
            bl20_position1 INTEGER,
            bl21_template2 INTEGER,
            bl22_edge2 REAL,
            bl23_corner2 INTEGER,
            bl24_position2 INTEGER,
            bl25_template3 INTEGER,
            bl26_edge3 REAL,
            bl27_corner3 INTEGER,
            bl28_position3 INTEGER,
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
            ORDER BY id
        ''', (file_id,))
        return cursor.fetchall()

    # 按文件名搜索
    @db_connection
    def search_by_filename(self, cursor, filename_pattern):
        """
        通过文件名模式搜索文件及其相关数据
        """
        try:
            query = '''
                SELECT g.*, f.directory, f.file_name
                FROM glass_data g
                JOIN file_info f ON g.file_id = f.id
                WHERE f.file_name LIKE LOWER(?)
                ORDER BY f.directory, g.layout_number, g.id
            '''
            cursor.execute(query, (f'%{filename_pattern}%',))
            results = cursor.fetchall()
            return results
        except Exception as e:
            print(f"搜索文件时出错: {e}")
            return []
      
    # 按分组号搜索
    @db_connection
    def search_by_group(self, cursor, group_number):
        """
        通过分组号搜索相关数据
        """
        try:
            group_number = str(group_number) if group_number is not None else ""
            if '-' in group_number:
                group_number, file_name = group_number.split('-')
                query = '''
                    SELECT g.*, f.directory, f.file_name
                    FROM glass_data g
                    JOIN file_info f ON g.file_id = f.id
                    WHERE LOWER(g.bl9_group) LIKE LOWER(?) and f.file_name LIKE LOWER(?)
                    ORDER BY f.directory, g.layout_number, g.id
                '''
                cursor.execute(query, (f'%{group_number}%', f'%{file_name}%'))
            else:
                query = '''
                    SELECT g.*, f.directory, f.file_name
                    FROM glass_data g
                    JOIN file_info f ON g.file_id = f.id
                    WHERE LOWER(g.bl9_group) LIKE LOWER(?)
                    ORDER BY f.directory, g.layout_number, g.id
                '''
                cursor.execute(query, (f'%{group_number}%',))
            results = cursor.fetchall()
            return results
        except Exception as e:
            print(f"搜索分组时出错: {e}")
            return []

    # 按尺寸搜索
    @db_connection
    def get_pieces_by_size(self, cursor, width=None, include_edge=False):
        print(f"搜索尺寸: {width}, 磨边位: {'是' if include_edge else '否'}")
        
        if width:
            width = str(width)
            has_x = 'x' in width.lower()

            if include_edge:  # 勾选磨边位
                if has_x:
                    # 带x的尺寸搜索切割尺寸X和Y
                    dimensions = width.lower().split('x')
                    query = '''
                        SELECT g.*, f.directory, f.file_name
                        FROM glass_data g
                        JOIN file_info f ON g.file_id = f.id
                        WHERE g.bl1_size_x = ? AND g.bl2_size_y = ?
                        ORDER BY f.directory, g.layout_number, g.id
                    '''
                    params = (dimensions[0].strip(), dimensions[1].strip())
                else:
                    # 单数字搜索切割尺寸x或者y
                    query = '''
                        SELECT g.*, f.directory, f.file_name
                        FROM glass_data g
                        JOIN file_info f ON g.file_id = f.id
                        WHERE g.bl1_size_x = ? OR g.bl2_size_y = ?
                        ORDER BY f.directory, g.layout_number, g.id
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
                        WHERE g.bl11_order_size LIKE ?
                        ORDER BY f.directory, g.layout_number, g.id
                    '''
                    params = (search_size,)
                else:
                    # 单数字搜索订单尺寸
                    search_size = f"%{width}%"
                    query = '''
                        SELECT g.*, f.directory, f.file_name
                        FROM glass_data g
                        JOIN file_info f ON g.file_id = f.id
                        WHERE g.bl11_order_size LIKE ?
                        ORDER BY f.directory, g.layout_number, g.id
                    '''
                    params = (search_size,)
        else:
            # 如果没有输入尺寸,返回前300条数据
            query = '''
                SELECT g.*, f.directory, f.file_name
                FROM glass_data g
                JOIN file_info f ON g.file_id = f.id
                ORDER BY f.directory, g.layout_number, g.id
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
            file_path = get_file_path(directory, file_name)
            file_md5 = calculate_file_md5(file_path)
            
            return self.insert_data_with_timestamp(directory, file_name, pieces_data, file_md5, None, None, None)
        except Exception as e:
            print(f"插入数据时出错: {e}")
            raise

    @db_connection
    def insert_data_with_timestamp(self, cursor, directory, file_name, pieces_data, file_md5, file_mtime, file_size, file_crc32):
        """插入数据并包含完整检查信息"""
        try:
            file_path = get_file_path(directory, file_name)
            
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
                        bl1_size_x, bl2_size_y, bl3_x_coordinate, bl4_y_coordinate,
                        bl5_display_length, bl6_display_width, bl7_customer,
                        bl8_position, bl9_group, bl10_order, bl11_order_size,
                        bl12, bl13, bl14, bl15_barcode, bl16_count,
                        bl17_template1, bl18_edge1, bl19_corner1, bl20_position1,
                        bl21_template2, bl22_edge2, bl23_corner2, bl24_position2,
                        bl25_template3, bl26_edge3, bl27_corner3, bl28_position3
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    file_id,
                    piece.get('raw_width'),
                    piece.get('raw_height'),
                    piece.get('material_code'),
                    piece.get('layout_number'),
                    piece.get('total_layouts'),
                    piece.get('thickness'),
                    piece.get('bl1_size_x'),
                    piece.get('bl2_size_y'),
                    piece.get('bl3_x_coordinate'),
                    piece.get('bl4_y_coordinate'),
                    piece.get('bl5_display_length'),
                    piece.get('bl6_display_width'),
                    piece.get('bl7_customer'),
                    piece.get('bl8_position'),
                    piece.get('bl9_group'),
                    piece.get('bl10_order'),
                    piece.get('bl11_order_size'),
                    piece.get('bl12'),
                    piece.get('bl13'),
                    piece.get('bl14'),
                    piece.get('bl15_barcode'),
                    piece.get('bl16_count'),
                    piece.get('bl17_template1'),
                    piece.get('bl18_edge1'),
                    piece.get('bl19_corner1'),
                    piece.get('bl20_position1'),
                    piece.get('bl21_template2'),
                    piece.get('bl22_edge2'),
                    piece.get('bl23_corner2'),
                    piece.get('bl24_position2'),
                    piece.get('bl25_template3'),
                    piece.get('bl26_edge3'),
                    piece.get('bl27_corner3'),
                    piece.get('bl28_position3')
                ))
            return True
        except Exception as e:
            print(f"插入数据时出错: {e}")
            raise

# Excel导出器
class ExcelExporter:
    def __init__(self):
        self.template_path = os.path.join(os.path.dirname(__file__), EXCEL_TEMPLATE)
        self.export_dir = EXPORT_DIRECTORY
        self.template_columns = {
            'A': '材料代码',
            'B': '成品长度',
            'C': '成品宽度',
            'D': '需切数量',
            'E': '客户',
            'F': '分组号',
            'G': '订单号',
            'H': '料号',
            'I': '位置号',
            'J': '条码', 
            'K': '订单尺寸', 
            'L': '基准边'
        }
        self.last_export_file = None
        self.last_material_code = None

    def get_export_filename(self, material_code, total_pieces=None):
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
        print(f"开始导出数据: material_code={material_code}, 总条数={len(data)}, append={append}")
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
                    self.material_mismatch = True
                    return False
                
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
                    # A列: 成品名称 = material_code (使用每行自己的material_code)
                    row_material_code = row.get('material_code', '') if row.get('material_code', '') else (material_code if material_code else '')
                    ws.write(current_row, 0, row_material_code)
                    # B列: 成品长度 = bl5_display_length (小片显示的长)
                    ws.write(current_row, 1, row.get('bl5_display_length', ''))
                    # C列: 成品宽度 = bl6_display_width (小片显示的宽)
                    ws.write(current_row, 2, row.get('bl6_display_width', ''))
                    # D列: 需切数量 = bl16_count
                    ws.write(current_row, 3, row.get('bl16_count', 1))
                    # E列: 客户 = bl7_customer
                    ws.write(current_row, 4, row.get('bl7_customer', ''))
                    # F列: 分组号 = bl9_group
                    ws.write(current_row, 5, row.get('bl9_group', ''))
                    # G列: 订单号 = bl10_order
                    ws.write(current_row, 6, row.get('bl10_order', ''))
                    # H列: 料号 = bl17_template1 (标签1激光模板序号)
                    template1 = row.get('bl17_template1', '')
                    ws.write(current_row, 7, int(template1) if template1 else '')
                    # I列: 位置号 = bl19_corner1 (标签1激光打码角位号)
                    corner1 = row.get('bl19_corner1', '')
                    ws.write(current_row, 8, int(corner1) if corner1 else '')
                    # J列: 条码 = bl15_barcode
                    ws.write(current_row, 9, row.get('bl15_barcode', ''))
                    # K列: 订单尺寸 = bl11_order_size
                    ws.write(current_row, 10, row.get('bl11_order_size', ''))
                    # L列: 基准边 = bl18_edge1
                    edge1 = row.get('bl18_edge1', '')
                    ws.write(current_row, 11, float(edge1) if edge1 else '')

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
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.parser = BaolunGCodeParser()
        self.db = DatabaseManager()
        self.exporter = ExcelExporter()
        self.default_dir = WORK_DIRECTORY
        
        # 添加状态栏
        self.statusBar()
        self.scan_status = QLabel()
        self.search_status = QLabel()
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
        self.statusBar().addPermanentWidget(self.db_update_status)
        # 更新数据库修改时间显示
        self.update_db_status()

        # 创建表格 - 根据带*号的字段显示
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            '目录', '文件名', '客户名称', '架号', '订单号', '订单尺寸', '条码号', '料号', '位置号'
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
        self.table.setColumnWidth(3, 100)  # 架号
        self.table.setColumnWidth(4, 120)  # 订单号
        self.table.setColumnWidth(5, 100)  # 订单尺寸
        self.table.setColumnWidth(6, 150)  # 条码号
        self.table.setColumnWidth(7, 80)   # 料号
        self.table.setColumnWidth(8, 60)   # 位置号
        
        # 启用隔行变色
        self.table.setAlternatingRowColors(True)
        
        # 设置表头属性
        header = self.table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header.setStretchLastSection(True)
        
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
            db_path = LOCAL_DB_NAME
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
        self.setWindowTitle('宝伦玻璃切割数据管理')
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowIcon(QIcon("./icon.ico"))
        # 创建主布局
        main_layout = QVBoxLayout()
        main_layout.setSpacing(10)
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # 创建顶部工具栏并固定
        toolbar = QToolBar()
        toolbar.setMovable(False)
        toolbar.setFloatable(False)
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
        search_widget.setFixedHeight(50)
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
        search_layout.setContentsMargins(10, 5, 10, 5)

        # 创建搜索类型单选按钮组
        self.search_type_group = QButtonGroup(self)
        
        # 创建单选按钮
        self.size_radio = QRadioButton("尺寸")
        self.filename_radio = QRadioButton("文件名")
        self.group_radio = QRadioButton("架号")
        self.size_radio.setContentsMargins(0, 0, 10, 0)
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

        try:
            # 获取数据库中所有文件信息（包括时间戳）
            conn = self.db.connect()
            cursor = conn.cursor()
            cursor.execute('SELECT directory, file_name, file_md5, last_processed_time FROM file_info')
            existing_files = {(row[0], row[1]): (row[2], row[3]) for row in cursor.fetchall()}
            conn.close()
            
            # 统计总文件数并收集需要处理的文件
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
            
            # 优化：先使用文件修改时间作为初步筛选
            need_process_files = []
            db_check_start = time.time()
            
            # 获取数据库中所有文件的修改时间
            file_time_cache = {}
            conn = self.db.connect()
            cursor = conn.cursor()
            cursor.execute("SELECT directory, file_name, file_md5, last_processed_time FROM file_info")
            for row in cursor.fetchall():
                file_time_cache[(row[0], row[1])] = (row[2], row[3])
            conn.close()
            
            # 使用智能策略进行快速筛选
            for file_path, file, directory in g_files:
                existing_info = file_time_cache.get((directory, file))
                
                if not existing_info:
                    # 新文件需要处理
                    need_process_files.append((file_path, file, directory, None))
                else:
                    existing_md5, last_processed_time = existing_info
                    
                    # 检查文件是否在一周内处理过
                    one_week_ago = time.time() - (7 * 24 * 3600)
                    
                    if last_processed_time and last_processed_time > one_week_ago:
                        # 一周内处理过的文件，可以跳过MD5计算（除非文件修改时间变化）
                        try:
                            file_mtime = os.path.getmtime(file_path)
                            need_process_files.append((file_path, file, directory, existing_md5))
                        except:
                            need_process_files.append((file_path, file, directory, existing_md5))
                    else:
                        # 超过一周未处理的文件，需要检查
                        need_process_files.append((file_path, file, directory, existing_md5))
            
            # 处理需要更新的文件
            print(f"开始处理 {len(need_process_files)} 个文件...")
            for current_count, (file_path, file, directory, existing_md5) in enumerate(need_process_files, 1):
                if progress.wasCanceled():
                    break
                    
                progress_value = int((current_count / len(need_process_files)) * 100)
                progress.setValue(progress_value)
                progress.setLabelText(f"正在处理文件 ({current_count}/{len(need_process_files)}): {file}")
                
                try:
                    # 优化：只有在数据库中存在记录时才计算MD5进行对比
                    if existing_md5 is not None:
                        current_md5 = calculate_file_md5(file_path)
                        
                        if not current_md5:
                            print(f"MD5计算失败，跳过文件: {file}")
                            continue
                        
                        # 检查文件是否需要处理
                        if existing_md5 == current_md5:
                            unchanged_files += 1
                            continue
                    else:
                        current_md5 = calculate_file_md5(file_path)
                    
                    # 解析G代码文件
                    pieces_data = self.parser.parse_file(file_path)
                    
                    if not pieces_data:
                        print(f"无法解析文件: {file}")
                        continue
                    
                    # 更新统计
                    if existing_md5 is None:
                        new_files += 1
                    else:
                        updated_files += 1
                    
                    total_pieces += len(pieces_data)
                    changed_pieces += len(pieces_data)
                    
                    # 插入数据库
                    self.db.insert_data(directory, file, pieces_data)
                    print(f"已处理文件: {file}, 小片数: {len(pieces_data)}")
                    
                except Exception as e:
                    print(f"处理文件 {file} 时出错: {e}")
                    continue
            
            progress.setValue(100)
            progress.close()
            
            status_text = f"扫描完成 - 总文件: {total_files} | 新增: {new_files} | 更新: {updated_files} | 未变化: {unchanged_files}"
            self.scan_status.setText(status_text)
            
            print(f"\n扫描完成:")
            print(f"总文件数: {total_files}")
            print(f"新文件: {new_files}")
            print(f"更新文件: {updated_files}")
            print(f"未变化文件: {unchanged_files}")
            print(f"总小片数: {total_pieces}")
            
        except Exception as e:
            print(f"处理目录时出错: {e}")
            if progress:
                progress.close()
            QMessageBox.critical(self, '错误', f'处理目录时出错: {e}')

    def load_data(self):
        """加载数据到表格"""
        try:
            conn = self.db.connect()
            cursor = conn.cursor()
            
            # 获取前300条数据
            query = '''
                SELECT g.*, f.directory, f.file_name
                FROM glass_data g
                JOIN file_info f ON g.file_id = f.id
                ORDER BY f.directory, g.layout_number, g.id
                LIMIT 300
            '''
            cursor.execute(query)
            results = cursor.fetchall()
            conn.close()
            
            self.display_results(results)
            
        except Exception as e:
            print(f"加载数据时出错: {e}")

    def search_data(self):
        """搜索数据"""
        try:
            search_input = self.search_input.text().strip()
            
            if not search_input:
                # 如果没有输入搜索内容，显示前300条数据
                self.load_data()
                self.search_status.setText("显示前300条数据")
                return
            
            conn = self.db.connect()
            cursor = conn.cursor()
            results = []
            
            # 根据选择的搜索类型执行不同的搜索
            if self.size_radio.isChecked():
                # 尺寸搜索
                include_edge = self.edge_checkbox.isChecked()
                results = self.db.get_pieces_by_size(search_input, include_edge)
                self.search_status.setText(f"尺寸 '{search_input}' - 找到 {len(results)} 条")
                
            elif self.filename_radio.isChecked():
                # 文件名搜索
                results = self.db.search_by_filename(search_input)
                self.search_status.setText(f"文件名 '{search_input}' - 找到 {len(results)} 条")
                
            elif self.group_radio.isChecked():
                # 架号搜索
                results = self.db.search_by_group(search_input)
                self.search_status.setText(f"架号 '{search_input}' - 找到 {len(results)} 条")
            
            conn.close()
            self.display_results(results)
            
        except Exception as e:
            print(f"搜索数据时出错: {e}")
            QMessageBox.critical(self, '错误', f'搜索数据时出错: {e}')

    def display_results(self, results):
        """显示搜索结果"""
        try:
            self.table.setRowCount(0)

            if not results:
                print("没有找到数据")
                return

            # 数据库字段索引映射
            # glass_data: id(0), file_id(1), raw_width(2), raw_height(3), material_code(4), layout_number(5), total_layouts(6), thickness(7)
            # bl1_size_x(8), bl2_size_y(9), bl3_x_coordinate(10), bl4_y_coordinate(11), bl5_display_length(12), bl6_display_width(13)
            # bl7_customer(14), bl8_position(15), bl9_group(16), bl10_order(17), bl11_order_size(18)
            # bl12(19), bl13(20), bl14(21), bl15_barcode(22), bl16_count(23)
            # bl17_template1(24), bl18_edge1(25), bl19_corner1(26), bl20_position1(27)
            # file_info: directory(37), file_name(38)

            for row_data in results:
                row_position = self.table.rowCount()
                self.table.insertRow(row_position)

                # 目录 - 存储glass_data的id在UserRole
                dir_item = QTableWidgetItem(str(row_data[36]))
                dir_item.setData(Qt.UserRole, row_data[0])  # 存储glass_data的id
                self.table.setItem(row_position, 0, dir_item)
                # 文件名
                self.table.setItem(row_position, 1, QTableWidgetItem(str(row_data[37])))
                # 客户名称 (bl7_customer)
                self.table.setItem(row_position, 2, QTableWidgetItem(str(row_data[14])))
                # 架号 (bl9_group)
                self.table.setItem(row_position, 3, QTableWidgetItem(str(row_data[16])))
                # 订单号 (bl10_order)
                self.table.setItem(row_position, 4, QTableWidgetItem(str(row_data[17])))
                # 订单尺寸 (bl11_order_size)
                self.table.setItem(row_position, 5, QTableWidgetItem(str(row_data[18])))
                # 条码号 (bl15_barcode)
                self.table.setItem(row_position, 6, QTableWidgetItem(str(row_data[22])))
                # 料号 (bl17_template1)
                self.table.setItem(row_position, 7, QTableWidgetItem(str(row_data[24])))
                # 位置号 (bl19_corner1)
                self.table.setItem(row_position, 8, QTableWidgetItem(str(row_data[26])))

            print(f"已显示 {len(results)} 条数据")

        except Exception as e:
            print(f"显示结果时出错: {e}")

    def export_excel(self, append=False):
        """导出Excel - 只导出选择的行,每行数量为1"""
        try:
            # 获取选中的行
            selected_rows = self.table.selectionModel().selectedRows()

            if not selected_rows:
                QMessageBox.warning(self, '提示', '请选择要导出的数据')
                return

            if len(selected_rows) > 300:
                reply = QMessageBox.question(
                    self,
                    '确认导出',
                    f'您选择了 {len(selected_rows)} 条数据，是否继续导出？',
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return

            # 从数据库获取选中行的完整数据
            data_to_export = []
            material_code = None

            conn = self.db.connect()
            cursor = conn.cursor()

            for index in sorted([row.row() for row in selected_rows]):
                try:
                    # 从表格获取glass_data的id
                    glass_data_id = self.table.item(index, 0).data(Qt.UserRole) if self.table.item(index, 0) else None

                    if not glass_data_id:
                        print(f"第 {index} 行没有glass_data_id")
                        continue

                    # 从数据库获取该行的glass_data
                    cursor.execute('''
                        SELECT
                            material_code,
                            bl5_display_length, bl6_display_width, bl7_customer,
                            bl9_group, bl10_order, bl11_order_size, bl15_barcode,
                            bl17_template1, bl18_edge1, bl19_corner1, bl16_count
                        FROM glass_data
                        WHERE id = ?
                    ''', (glass_data_id,))

                    piece = cursor.fetchone()
                    if piece:
                        # 从记录获取material_code
                        if not material_code and piece[0]:
                            material_code = piece[0]

                        row_data = {
                            'material_code': piece[0] if piece[0] else '',
                            'bl5_display_length': piece[1] if piece[1] else 0,
                            'bl6_display_width': piece[2] if piece[2] else 0,
                            'bl7_customer': piece[3] if piece[3] else '',
                            'bl9_group': piece[4] if piece[4] else '',
                            'bl10_order': piece[5] if piece[5] else '',
                            'bl11_order_size': piece[6] if piece[6] else '',
                            'bl15_barcode': piece[7] if piece[7] else '',
                            'bl17_template1': piece[8] if piece[8] else '',
                            'bl18_edge1': piece[9] if piece[9] else 0,
                            'bl19_corner1': piece[10] if piece[10] else '',
                            'bl16_count': 1  # 每行数量固定为1
                        }
                        data_to_export.append(row_data)
                    else:
                        print(f"第 {index} 行未找到glass_data记录")

                except Exception as e:
                    print(f"处理第 {index} 行数据时出错: {e}")
                    continue

            conn.close()

            if not data_to_export:
                QMessageBox.warning(self, '提示', '没有可导出的数据')
                return

            if not material_code:
                material_code = 'unknown'

            # 导出Excel
            success = self.exporter.export_data(data_to_export, material_code, len(data_to_export), append)

            if success:
                action = "追加导出" if append else "导出"
                QMessageBox.information(self, '成功', f'{action}完成！共导出 {len(data_to_export)} 条数据')
            else:
                QMessageBox.warning(self, '失败', '导出失败，请检查错误信息')

        except Exception as e:
            print(f"导出Excel时出错: {e}")
            QMessageBox.critical(self, '错误', f'导出Excel时出错: {e}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
