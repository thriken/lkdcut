import os
import pandas as pd
import hashlib
import sqlite3
from datetime import datetime

class ExcelMonitor:
    def __init__(self):
        self.backup_dir = r'\\landierp\办公室\补片文件备份'
        self.db_file = 'file_history.db'
        self.target_columns = ['A', 'B', 'C', 'D', 'R', 'U', 'V', 'W', 'X', 'Y']
        self.column_mappings = {
            'A': '产品类别(材料)',
            'B': '宽',
            'C': '高',
            'D': '数量',
            'R': '条码号(注释5)',
            'U': '注释7订单尺寸',
            'V': '基准线(注释8)',
            'W': '注释9架号',
            'X': '第一标位置(注释10)',
            'Y': '第二标(注释11)'
        }
        self.init_db()

    def init_db(self):
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            # 创建文件表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS files (
                    file_path TEXT PRIMARY KEY,
                    file_hash TEXT,
                    last_modified TEXT
                )
            ''')
            # 创建数据表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    file_path TEXT,
                    col_a TEXT, col_b TEXT, col_c TEXT, col_d TEXT,
                    col_r TEXT, col_u TEXT, col_v TEXT, col_w TEXT,
                    col_x TEXT, col_y TEXT,
                    FOREIGN KEY (file_path) REFERENCES files (file_path)
                )
            ''')
            conn.commit()

    def get_file_info(self, file_path):
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT file_hash FROM files WHERE file_path = ?', (file_path,))
            result = cursor.fetchone()
            return result[0] if result else None

    def save_file_data(self, file_path, file_hash, data):
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            # 保存文件信息
            cursor.execute('''
                INSERT OR REPLACE INTO files (file_path, file_hash, last_modified)
                VALUES (?, ?, ?)
            ''', (file_path, file_hash, datetime.now().isoformat()))
            
            # 删除旧数据
            cursor.execute('DELETE FROM data WHERE file_path = ?', (file_path,))
            
            # 保存新数据
            if data:
                for row in data:
                    cursor.execute('''
                        INSERT INTO data (
                            file_path, col_a, col_b, col_c, col_d,
                            col_r, col_u, col_v, col_w, col_x, col_y
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        file_path,
                        str(row.get('A', '')), str(row.get('B', '')),
                        str(row.get('C', '')), str(row.get('D', '')),
                        str(row.get('R', '')), str(row.get('U', '')),
                        str(row.get('V', '')), str(row.get('W', '')),
                        str(row.get('X', '')), str(row.get('Y', ''))
                    ))
            conn.commit()

    def delete_file_data(self, file_path):
        with sqlite3.connect(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM data WHERE file_path = ?', (file_path,))
            cursor.execute('DELETE FROM files WHERE file_path = ?', (file_path,))
            conn.commit()

    def calculate_file_hash(self, file_path):
        with open(file_path, 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()

    def process_excel(self, file_path):
        try:
            if not os.path.exists(file_path):
                return None
            # 根据文件扩展名选择合适的引擎
            engine = 'xlrd' if file_path.endswith('.xls') else 'openpyxl'

            # 读取Excel文件，跳过第一行标题
            df = pd.read_excel(
                file_path,
                engine=engine,
                header=None,
                skiprows=1,
                dtype=str  # 将所有列读取为字符串类型
            )
            
            # 检查是否有足够的列
            if len(df.columns) < 4:  # 至少需要产品类别、宽、高、数量四列
                return None
            
            # 创建一个新的DataFrame，包含所有需要的列
            result_df = pd.DataFrame(columns=self.target_columns)
            
            # 定义列映射关系
            column_mapping = {
                'A': 0,  # 产品类别(材料)
                'B': 1,  # 宽
                'C': 2,  # 高
                'D': 3,  # 数量
                'R': 17,  # 条码号
                'U': 20,  # 订单尺寸
                'V': 21,  # 基准线
                'W': 22,  # 架号
                'X': 23,  # 第一标位置
                'Y': 24   # 第二标
            }
            
            # 根据映射关系提取数据
            for target_col, source_col in column_mapping.items():
                if source_col < len(df.columns):
                    result_df[target_col] = df[source_col].fillna('')
                else:
                    result_df[target_col] = ''
            
            # 数据清理和验证
            result_df = result_df.dropna(subset=['A', 'B', 'C', 'D'], how='all')
            
            # 确保数值列为数值类型
            for col in ['B', 'C', 'D']:
                result_df[col] = pd.to_numeric(result_df[col], errors='coerce')
            
            # 移除无效数据的行
            result_df = result_df.dropna(subset=['B', 'C', 'D'])
            
            # 将数值转换回字符串以保持一致性
            for col in result_df.columns:
                result_df[col] = result_df[col].astype(str)
            
            # 将DataFrame转换为字典列表
            return result_df.to_dict('records')
        except Exception as e:
            return None

    def scan_files(self):
        changes = {'new': [], 'modified': [], 'removed': []}
        current_files = set()

        try:
            # 验证网络路径可访问性
            if not os.path.exists(self.backup_dir):
                return changes

            # 尝试列出目录内容以验证访问权限
            try:
                next(os.scandir(self.backup_dir))
            except StopIteration:
                return changes
            except PermissionError:
                return changes
            except Exception:
                return changes

            for root, _, files in os.walk(self.backup_dir):
                for file in files:
                    if file.endswith('.xls'):
                        file_path = os.path.join(root, file)
                        current_files.add(file_path)
                        try:
                            file_hash = self.calculate_file_hash(file_path)
                            stored_hash = self.get_file_info(file_path)
                            if not stored_hash:
                                changes['new'].append(file_path)
                                self.save_file_data(file_path, file_hash, self.process_excel(file_path))
                            elif stored_hash != file_hash:
                                changes['modified'].append(file_path)
                                self.save_file_data(file_path, file_hash, self.process_excel(file_path))
                        except Exception:
                            continue

            # 检查删除的文件
            try:
                with sqlite3.connect(self.db_file) as conn:
                    cursor = conn.cursor()
                    cursor.execute('SELECT file_path FROM files')
                    stored_files = {row[0] for row in cursor.fetchall()}
                    
                    for file_path in stored_files:
                        if file_path not in current_files:
                            changes['removed'].append(file_path)
                            self.delete_file_data(file_path)
            except Exception:
                pass

        except Exception:
            pass

        return changes

    def run(self):
        return self.scan_files()

if __name__ == '__main__':
    monitor = ExcelMonitor()
    monitor.run()