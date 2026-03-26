# -*- coding: utf-8 -*-
"""
图片水平翻转转换程序
规则：
1. 1-99 原图片 -> 101-199 镜像图片
2. 201-299 原图片 -> 301-399 镜像图片
"""
import os
from PIL import Image

def convert_images():
    img_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(img_dir, 'resources\img')

    if not os.path.exists(output_dir):
        print(f"目录不存在: {output_dir}")
        return

    # 获取所有png文件
    png_files = [f for f in os.listdir(output_dir) if f.endswith('.png')]

    converted_count = 0

    for filename in png_files:
        try:
            # 提取数字部分
            name_without_ext = os.path.splitext(filename)[0]
            if not name_without_ext.isdigit():
                continue

            num = int(name_without_ext)

            # 判断是否在需要转换的范围内
            # 1-99 -> 生成 101-199
            if 1 <= num <= 99:
                new_num = num + 100
            # 201-299 -> 生成 301-399
            elif 201 <= num <= 299:
                new_num = num + 100
            else:
                continue

            # 检查目标文件是否已存在
            new_filename = f"{new_num}.png"
            if os.path.exists(os.path.join(output_dir, new_filename)):
                print(f"跳过 {filename} -> {new_filename} (已存在)")
                continue

            # 打开图片并水平翻转
            img_path = os.path.join(output_dir, filename)
            img = Image.open(img_path)

            # 水平翻转 (左右翻转)
            flipped_img = img.transpose(Image.FLIP_LEFT_RIGHT)

            # 保存
            new_path = os.path.join(output_dir, new_filename)
            flipped_img.save(new_path)

            print(f"转换成功: {filename} -> {new_filename}")
            converted_count += 1

        except Exception as e:
            print(f"转换失败 {filename}: {e}")

    print(f"\n完成！共转换 {converted_count} 张图片")

if __name__ == '__main__':
    convert_images()
