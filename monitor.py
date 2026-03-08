import os
import re
import sqlite3
import hashlib
import time
import binascii
from datetime import datetime, timedelta

# Utils 功能
def get_current_time():
    """获取当前时间格式 hh:mm:ss"""
    return datetime.now().strftime("%H:%M:%S")

def log_message(message):
    """输出带时间戳的日志信息"""
    timestamp = get_current_time()
    print(f"[{timestamp}] {message}")

def get_file_path(root_dir, directory, file_name):
    """统一获取文件路径"""
    if not all([file_name]):  # 只检查文件名
        raise ValueError("文件名不能为空")
    
    # 使用固定的网络路径
    root_dir = r'\\landierp\Share\激光文件'
    
    if not directory or directory == '.':
        return os.path.join(root_dir, file_name)
    
    return os.path.join(root_dir, directory, file_name)

def calculate_file_md5(file_path):
    """计算文件MD5值"""
    try:
        file_path = os.path.normpath(file_path)
        if not os.path.exists(file_path):
            log_message(f"文件不存在: {file_path}")
            return None
            
        try:
            with open(file_path, 'rb') as f:
                md5_hash = hashlib.md5()
                for chunk in iter(lambda: f.read(4096), b""):
                    md5_hash.update(chunk)
                return md5_hash.hexdigest()
        except PermissionError:
            log_message(f"没有权限访问文件: {file_path}")
            return None
        except Exception as e:
            log_message(f"读取文件时出错: {file_path}, 错误: {e}")
            return None
    except Exception as e:
        log_message(f"计算MD5时出错: {e}")
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
        self.init_database()

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
                WHERE f.file_name LIKE ?
                ORDER BY f.directory, g.layout_number, g.piece_number
            '''
            cursor.execute(query, (f'%{filename_pattern}%',))
            results = cursor.fetchall()
            return results
        except Exception as e:
            log_message(f"搜索文件时出错: {e}")
            return []

    # 新增的搜索方法 按分组号搜索
    @db_connection
    def search_by_group(self, cursor, group_number):
        """
        通过分组号搜索相关数据
        :param cursor: 数据库游标
        :param group_number: 分组号
        :return: 搜索结果列表
        """
        try:
            query = '''
                SELECT g.*, f.directory, f.file_name
                FROM glass_data g
                JOIN file_info f ON g.file_id = f.id
                WHERE g.group_number = ?
                ORDER BY f.directory, g.layout_number, g.piece_number
            '''
            cursor.execute(query, (group_number,))
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
    def delete_folder_data(self, cursor, directory):
        """删除指定文件夹的所有数据"""
        try:
            # 删除glass_data表中该文件夹的文件记录
            cursor.execute('''
                DELETE FROM glass_data 
                WHERE file_id IN (
                    SELECT id FROM file_info 
                    WHERE directory = ?
                )
            ''', (directory,))
            
            # 删除file_info表中该文件夹的所有记录
            cursor.execute('''
                DELETE FROM file_info WHERE directory = ?
            ''', (directory,))
            
            log_message(f"已删除文件夹 {directory} 的所有数据")
            return True
        except Exception as e:
            log_message(f"删除文件夹数据时出错: {e}")
            raise

    @db_connection
    def insert_data_with_timestamp(self, cursor, directory, file_name, pieces_data, file_md5, file_mtime, file_size, file_crc32):
        """插入数据并包含完整检查信息"""
        try:
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
            log_message(f"插入数据时出错: {e}")
            raise

    @db_connection
    def delete_folder_data(self, cursor, directory):
        """删除指定文件夹的所有数据"""
        try:
            # 删除glass_data表中该文件夹的文件记录
            cursor.execute('''
                DELETE FROM glass_data 
                WHERE file_id IN (
                    SELECT id FROM file_info 
                    WHERE directory = ?
                )
            ''', (directory,))
            
            # 删除file_info表中该文件夹的所有记录
            cursor.execute('''
                DELETE FROM file_info WHERE directory = ?
            ''', (directory,))
            
            log_message(f"已删除文件夹 {directory} 的所有数据")
            return True
        except Exception as e:
            log_message(f"删除文件夹数据时出错: {e}")
            raise

    @db_connection
    def insert_data_with_timestamp(self, cursor, directory, file_name, pieces_data, file_md5, file_mtime, file_size, file_crc32):
        """插入数据并包含完整检查信息"""
        try:
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
            log_message(f"插入数据时出错: {e}")
            raise

    @db_connection
    def insert_data(self, cursor, directory, file_name, pieces_data):
        try:
            # 计算文件MD5
            file_path = get_file_path(os.path.dirname(__file__), directory, file_name)
            file_md5 = calculate_file_md5(file_path)
            
            # 删除旧数据
            cursor.execute('''
                DELETE FROM glass_data 
                WHERE file_id IN (
                    SELECT id FROM file_info 
                    WHERE directory = ? AND file_name = ?
                )
            ''', (directory, file_name))
            
            cursor.execute('''
                DELETE FROM file_info 
                WHERE directory = ? AND file_name = ?
            ''', (directory, file_name))
            
            # 插入文件信息
            cursor.execute('''
                INSERT INTO file_info (directory, file_name, file_md5)
                VALUES (?, ?, ?)
            ''', (directory, file_name, file_md5))
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


class Processer:
    def __init__(self):
        self.db = DatabaseManager()
        self.parser = GCodeParser()

    def get_files_from_last_week(self):
        """获取一周内被处理过的文件信息"""
        conn = self.db.connect()
        cursor = conn.cursor()
        
        # 计算一周前的时间戳
        one_week_ago = time.time() - (7 * 24 * 60 * 60)
        
        cursor.execute('''
            SELECT directory, file_name, file_md5, file_mtime, file_size, file_crc32 
            FROM file_info 
            WHERE last_processed_time >= ?
        ''', (one_week_ago,))
        
        result = {}
        for row in cursor.fetchall():
            directory, file_name, file_md5, file_mtime, file_size, file_crc32 = row
            result[(directory, file_name)] = {
                'md5': file_md5,
                'mtime': file_mtime,
                'size': file_size,
                'crc32': file_crc32
            }
        
        conn.close()
        return result

    def check_file_changes(self, file_path, existing_info):
        """检查文件是否发生变化"""
        if not os.path.exists(file_path):
            return 'deleted'
        
        current_mtime = os.path.getmtime(file_path)
        current_size = os.path.getsize(file_path)
        
        # 首先检查文件大小和时间戳
        if existing_info:
            if current_size != existing_info.get('size', 0) or abs(current_mtime - existing_info.get('mtime', 0)) > 1:
                # 文件大小或修改时间发生变化，需要进一步检查
                current_md5 = calculate_file_md5(file_path)
                if current_md5 != existing_info.get('md5'):
                    return 'changed'
        
        return 'unchanged'

    def process_directory(self, workdir):
        log_message(f"开始监控目录: {workdir}")
        total_files = 0
        changed_files = 0
        new_files = 0
        unchanged_files = 0
        processed_folders = set()

        try:
            log_message("正在统计文件数量...")
            
            # 获取一周内处理过的文件信息
            recent_files = self.get_files_from_last_week()
            log_message(f"发现一周内处理过的文件: {len(recent_files)} 个")
            
            # 收集需要处理的文件
            g_files = []
            folder_files = {}  # 按文件夹分组文件
            
            for root, _, files in os.walk(workdir):
                for file in files:
                    if file.endswith('.g'):
                        relative_dir = os.path.relpath(root, workdir)
                        if relative_dir == '.':
                            relative_dir = ''
                        
                        file_path = os.path.join(root, file)
                        
                        # 只检查一周内处理过的文件或者新文件
                        existing_info = recent_files.get((relative_dir, file))
                        file_status = self.check_file_changes(file_path, existing_info)
                        
                        total_files += 1
                        
                        if file_status == 'changed':
                            changed_files += 1
                            g_files.append((file_path, file, relative_dir))
                            # 记录文件夹
                            if relative_dir not in folder_files:
                                folder_files[relative_dir] = []
                            folder_files[relative_dir].append(file)
                        elif file_status == 'deleted':
                            changed_files += 1
                            # 记录文件夹
                            if relative_dir not in folder_files:
                                folder_files[relative_dir] = []
                            folder_files[relative_dir].append(file)
                        elif not existing_info:  # 新文件
                            new_files += 1
                            g_files.append((file_path, file, relative_dir))
                            if relative_dir not in folder_files:
                                folder_files[relative_dir] = []
                            folder_files[relative_dir].append(file)
                        else:
                            unchanged_files += 1

            log_message(f"监控完成 - 总文件: {total_files} | 新增: {new_files} | 更新: {changed_files} | 未变化: {unchanged_files}")
            
            # 如果没有变化的文件，直接返回
            if not g_files:
                return

            # 按文件夹处理文件
            log_message("开始处理变化的文件...")
            processed_folders_count = 0
            
            for folder, files_in_folder in folder_files.items():
                processed_folders_count += 1
                log_message(f"处理文件夹 [{processed_folders_count}/{len(folder_files)}]: {folder if folder else '根目录'}")
                
                # 检查文件夹是否需要全量更新
                folder_needs_full_update = False
                folder_files_data = []
                
                for file in files_in_folder:
                    file_path = os.path.join(workdir, folder, file) if folder else os.path.join(workdir, file)
                    
                    if not os.path.exists(file_path):
                        # 文件被删除，需要全量更新文件夹
                        folder_needs_full_update = True
                        log_message(f"发现文件被删除: {file}")
                        break
                    
                    # 计算文件完整信息
                    file_md5 = calculate_file_md5(file_path)
                    file_mtime = os.path.getmtime(file_path)
                    file_size = os.path.getsize(file_path)
                    file_crc32 = calculate_file_crc32(file_path)
                    
                    if not file_md5:
                        log_message(f"跳过文件（MD5计算失败）: {file}")
                        continue
                    
                    pieces_data = self.parser.parse_file(file_path)
                    if not pieces_data:
                        log_message(f"文件解析失败，跳过: {file}")
                        continue
                    
                    folder_files_data.append({
                        'file': file,
                        'path': file_path,
                        'md5': file_md5,
                        'mtime': file_mtime,
                        'size': file_size,
                        'crc32': file_crc32,
                        'pieces': pieces_data
                    })

                # 如果文件夹中有文件被删除，或者有文件MD5变化，需要全量更新整个文件夹
                if folder_needs_full_update or len(folder_files_data) < len(files_in_folder):
                    # 删除整个文件夹的数据
                    self.db.delete_folder_data(folder)
                    log_message(f"删除文件夹 {folder} 的所有数据")
                    
                    # 重新插入所有存在的文件
                    for file_data in folder_files_data:
                        try:
                            if self.db.insert_data_with_timestamp(
                                folder, file_data['file'], file_data['pieces'],
                                file_data['md5'], file_data['mtime'], file_data['size'], file_data['crc32']
                            ):
                                log_message(f"插入文件数据成功: {file_data['file']} ({len(file_data['pieces'])}个小片)")
                        except Exception as e:
                            log_message(f"处理文件失败: {file_data['file']}, 错误: {e}")
                else:
                    # 只更新变化的文件
                    for file_data in folder_files_data:
                        try:
                            if self.db.insert_data_with_timestamp(
                                folder, file_data['file'], file_data['pieces'],
                                file_data['md5'], file_data['mtime'], file_data['size'], file_data['crc32']
                            ):
                                log_message(f"更新文件数据成功: {file_data['file']} ({len(file_data['pieces'])}个小片)")
                        except Exception as e:
                            log_message(f"处理文件失败: {file_data['file']}, 错误: {e}")

            log_message(f"处理完成! 共处理 {len(folder_files)} 个文件夹")
            
        except Exception as e:
            log_message(f"处理目录时出错: {e}")


def main():
    try:
        processor = Processer()
        workdir = r'\\landierp\Share\激光文件'
        processor.process_directory(workdir)
    except Exception as e:
        log_message(f"程序运行出错: {e}")
        input("按回车键退出...")

if __name__ == "__main__":
    main()
