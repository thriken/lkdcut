#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
将3c.csv转换为SQLite数据库
"""

import csv
import sqlite3
import os


def convert_csv_to_sqlite():
    """转换CSV到SQLite数据库"""

    csv_file = os.path.join(os.path.dirname(__file__), '3c.csv')
    db_file = os.path.join(os.path.dirname(__file__), '3c.db')

    if not os.path.exists(csv_file):
        print(f"CSV文件不存在: {csv_file}")
        return

    # 尝试不同的编码读取CSV
    encodings = ['utf-8-sig', 'gb2312', 'gbk', 'utf-8']
    rows = None

    for encoding in encodings:
        try:
            with open(csv_file, 'r', encoding=encoding) as f:
                reader = csv.reader(f)
                rows = list(reader)
                print(f"成功使用编码 {encoding} 读取CSV文件")
                break
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f"使用编码 {encoding} 读取失败: {e}")
            continue

    if rows is None:
        print("无法读取CSV文件")
        return

    # 删除旧数据库
    if os.path.exists(db_file):
        os.remove(db_file)

    # 创建数据库连接
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # 创建表
    cursor.execute('''
        CREATE TABLE materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            process_num TEXT,
            project_name TEXT,
            omron_laser TEXT,
            jingling_standard TEXT,
            reverse_standard TEXT,
            lowe_special TEXT,
            remark TEXT,
            image_path TEXT,
            lowe_image_path TEXT,
            material_info TEXT
        )
    ''')

    # 创建索引
    cursor.execute('CREATE INDEX idx_jingling ON materials(jingling_standard)')
    cursor.execute('CREATE INDEX idx_reverse ON materials(reverse_standard)')
    cursor.execute('CREATE INDEX idx_lowe ON materials(lowe_special)')

    # 插入数据（跳过标题行）
    inserted_count = 0
    for idx, row in enumerate(rows[1:], start=2):  # start=2 因为第1行是标题，第2行是数据
        # 确保行有足够的数据（至少9列：0-8）
        if len(row) >= 9:
            try:
                # 填充缺失的列
                while len(row) < 9:
                    row.append('')

                # 将料号数据转换为JSON字符串，存储所有料号及其类型
                import json
                material_info = {}

                # 打印前几行的详细数据
                if idx <= 10:
                    print(f"\n[DEBUG] 第{idx}行:")
                    print(f"  列0(加工网版号): {row[0]}")
                    print(f"  列1(项目名): {row[1]}")
                    print(f"  列2(欧姆激光标): {row[2]}")
                    print(f"  列3(精菱正标料号): {row[3]}")
                    print(f"  列4(反标料号): {row[4]}")
                    print(f"  列5(专用LOWE料号): {row[5]}")
                    print(f"  列6(备注): {row[6]}")
                    print(f"  列7(示例图片): {row[7]}")
                    print(f"  列8(LOW专用): {row[8]}")

                # 正标料号
                if row[3] and row[3].strip():
                    material_info[row[3].strip()] = {
                        'type': '正标',
                        'image': row[7].strip() if len(row) > 7 and row[7] else ''
                    }
                    if idx <= 10:
                        print(f"  -> 添加正标料号: {row[3].strip()}, 图片: {row[7] if len(row) > 7 else ''}")

                # 反标料号
                if row[4] and row[4].strip():
                    material_info[row[4].strip()] = {
                        'type': '反标',
                        'image': row[7].strip() if len(row) > 7 and row[7] else ''
                    }
                    if idx <= 10:
                        print(f"  -> 添加反标料号: {row[4].strip()}, 图片: {row[7] if len(row) > 7 else ''}")

                # LOWE专用料号
                if row[5] and row[5].strip():
                    material_info[row[5].strip()] = {
                        'type': 'LOWE专用标',
                        'image': row[8].strip() if len(row) > 8 and row[8] else ''
                    }
                    if idx <= 10:
                        print(f"  -> 添加LOWE专用标料号: {row[5].strip()}, 图片: {row[8] if len(row) > 8 else ''}")

                # 存储JSON
                material_json = json.dumps(material_info, ensure_ascii=False)
                if idx <= 10:
                    print(f"  -> JSON: {material_json}")

                cursor.execute('''
                    INSERT INTO materials (
                        process_num, project_name, omron_laser,
                        jingling_standard, reverse_standard, lowe_special,
                        remark, image_path, lowe_image_path, material_info
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    row[0].strip() if row[0] else '',
                    row[1].strip() if row[1] else '',
                    row[2].strip() if row[2] else '',
                    row[3].strip() if row[3] else '',
                    row[4].strip() if row[4] else '',
                    row[5].strip() if row[5] else '',
                    row[6].strip() if row[6] else '',
                    row[7].strip() if row[7] else '',
                    row[8].strip() if row[8] else '',
                    material_json
                ))
                inserted_count += 1
            except Exception as e:
                print(f"插入数据失败: {e}, 行内容: {row}")

    conn.commit()
    conn.close()

    print(f"\n转换完成！")
    print(f"CSV文件: {csv_file}")
    print(f"数据库文件: {db_file}")
    print(f"共插入 {inserted_count} 条记录")


if __name__ == '__main__':
    convert_csv_to_sqlite()
