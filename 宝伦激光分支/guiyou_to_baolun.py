#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
贵友G文件转换为宝伦G文件
"""

import os
import sys
import re


def parse_guiyou_data(line):
    """解析贵友P400x行的数据"""
    parts = line.strip().split('_')
    
    # 贵友字段映射 (16个字段)
    # 1-6: 通用尺寸坐标
    # 7: 客户名
    # 8: 定位号
    # 9: 分组号
    # 10: 订单尺寸
    # 11: 条码
    # 12: 料号
    # 13: 位置号
    # 14: 切片数量
    # 15: 基准边
    
    guiyou_data = {
        'gy1': parts[0] if len(parts) > 0 else '',      # X方向尺寸
        'gy2': parts[1] if len(parts) > 1 else '',      # Y方向尺寸
        'gy3': parts[2] if len(parts) > 2 else '',      # 左下角X坐标
        'gy4': parts[3] if len(parts) > 3 else '',      # 左下角Y坐标
        'gy5': parts[4] if len(parts) > 4 else '',      # 显示的长(X)
        'gy6': parts[5] if len(parts) > 5 else '',      # 显示的宽(Y)
        'gy7': parts[6] if len(parts) > 6 else '',      # 客户名
        'gy8': parts[7] if len(parts) > 7 else '',      # 定位号
        'gy9': parts[8] if len(parts) > 8 else '',      # 分组号
        'gy10': parts[9] if len(parts) > 9 else '',     # 订单尺寸
        'gy11': parts[10] if len(parts) > 10 else '',   # 条码
        'gy12': parts[11] if len(parts) > 11 else '',   # 料号
        'gy13': parts[12] if len(parts) > 12 else '',   # 位置号
        'gy14': parts[13] if len(parts) > 13 else '',   # 切片数量
        'gy15': parts[14] if len(parts) > 14 else '',   # 基准边
    }
    
    return guiyou_data


def convert_to_baolun(guiyou_data, piece_num):
    """将贵友数据转换为宝伦格式 (28个字段)"""
    # 宝伦字段映射 (参考贵友转宝伦.md)
    baolun_parts = [
        guiyou_data['gy1'],          # 宝伦1: X方向尺寸
        guiyou_data['gy2'],          # 宝伦2: Y方向尺寸
        guiyou_data['gy3'],          # 宝伦3: 左下角X坐标
        guiyou_data['gy4'],          # 宝伦4: 左下角Y坐标
        guiyou_data['gy5'],          # 宝伦5: 显示的长(X)
        guiyou_data['gy6'],          # 宝伦6: 显示的宽(Y)
        guiyou_data['gy7'],          # 宝伦7: 客户名
        guiyou_data['gy8'],          # 宝伦8: 定位号
        guiyou_data['gy9'],          # 宝伦9: 货架号分组号
        '',                          # 宝伦10: 订单号 (贵友没有对应,留空)
        guiyou_data['gy10'],         # 宝伦11: 订单尺寸加片标记
        '',                          # 宝伦12: 空
        '',                          # 宝伦13: 空
        '',                          # 宝伦14: 空
        guiyou_data['gy11'],         # 宝伦15: 条码号
        guiyou_data['gy14'],         # 宝伦16: 小片数量
        guiyou_data['gy12'],         # 宝伦17: 标签1激光模板序号
        guiyou_data['gy15'],         # 宝伦18: 标签1基准边
        guiyou_data['gy13'],         # 宝伦19: 标签1角位号
        '1',                         # 宝伦20: 固定1
        '3',                         # 宝伦21: 固定3
        guiyou_data['gy15'],         # 宝伦22: 标签2基准边
        guiyou_data['gy13'],         # 宝伦23: 标签2角位号
        '1',                         # 宝伦24: 固定1
        '',                          # 宝伦25: 未使用
        '',                          # 宝伦26: 未使用
        '',                          # 宝伦27: 未使用
        '',                          # 宝伦28: 未使用
    ]
    
    return '_'.join(baolun_parts)


def convert_guiyou_to_baolun(input_file, output_file=None):
    """转换贵友G文件为宝伦G文件"""
    try:
        # 读取文件 - 优先使用 GB2312 编码
        try:
            with open(input_file, 'r', encoding='gb2312') as f:
                lines = f.readlines()
        except UnicodeDecodeError:
            try:
                with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = f.readlines()
            except Exception:
                print(f"无法读取文件: {input_file}")
                return None
        
        # 设置输出文件名
        if not output_file:
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            output_file = os.path.join(os.path.dirname(input_file), f"{base_name}_宝伦.G")
        
        # 解析并转换
        output_lines = []
        piece_count = 0
        
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 保留头部的通用参数 (N01-N12)
            if line.startswith('N') and re.match(r'^N\d{2}\s+P30\d{2}\s*=', line):
                # 跳过空的材料代码 (P3007 = )
                if 'P3007 =  ' not in line:
                    output_lines.append(line)

            # 处理P400x行 (小片数据) - 贵友格式可能用P4001-P4100
            match = re.match(r'^N\d{2}\s+P4(\d+)\s*=\s*(.+)', line)
            if match and match.group(1).isdigit():
                p_num = int(match.group(1))
                value = match.group(2)

                # 贵友的P400x行都需要转换
                if 1 <= p_num <= 100:  # P4001-P4100
                    # 解析贵友数据
                    guiyou_data = parse_guiyou_data(value)
                    piece_count += 1

                    # 转换为宝伦格式
                    baolun_value = convert_to_baolun(guiyou_data, piece_count)

                    # 生成宝伦格式的行 (使用N13-N16, P4001-P4007)
                    new_n_num = 12 + ((piece_count - 1) % 4) + 1  # N13-N16循环
                    new_p_num = 4000 + piece_count  # P4001-P4007
                    output_lines.append(f"N{new_n_num:02d}  P{new_p_num}= {baolun_value}")
                # 不保留原始的P400x行

            # 保留G代码部分
            elif line.startswith('G') or line.startswith('M'):
                output_lines.append(line)
        
        # 写入输出文件
        with open(output_file, 'w', encoding='gb2312') as f:
            f.write('\n'.join(output_lines) + '\n')
        
        print(f"转换成功!")
        print(f"输入文件: {input_file}")
        print(f"输出文件: {output_file}")
        print(f"共转换 {piece_count} 个小片")
        
        return output_file
        
    except Exception as e:
        print(f"转换失败: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使用方法:")
        print("  python guiyou_to_baolun.py <贵友G文件路径> [输出文件路径]")
        print("\n示例:")
        print("  python guiyou_to_baolun.py GY2323.G")
        print("  python guiyou_to_baolun.py GY2323.G 输出_宝伦.G")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(input_file):
        print(f"错误: 文件不存在 - {input_file}")
        sys.exit(1)
    
    convert_guiyou_to_baolun(input_file, output_file)
