#!/usr/bin/env python3
"""
打包工具 — 生成 Windows 也能正常解压的 zip
Mac 自带的"压缩"功能会产生 __MACOSX 目录和编码问题，
用这个脚本打包出来的 zip 纯净、中文不乱码、Windows 友好。

使用：python3 pack_zip.py    （或双击运行 run.command / run.bat 打包后再分享）
"""

import zipfile
import sys
from pathlib import Path

# 要打包的文件夹名称
FOLDER_NAME = 'fund_calc'

# 打包时要排除的文件/目录（Mac 垃圾 + 打包产物本身）
EXCLUDE_NAMES = {
    '.DS_Store',
    '__MACOSX',
    '.Thumbs.db',
    '._',          # Mac 资源分叉前缀
    '.AppleDouble',
}


def should_exclude(path: Path) -> bool:
    """判断文件/目录是否应该跳过"""
    name = path.name
    # 排除隐藏文件（. 开头）
    if name.startswith('.'):
        return True
    # 排除已知垃圾名称
    if name in EXCLUDE_NAMES:
        return True
    # 排除文件预览图等
    if name.startswith('~$'):
        return True
    return False


def create_clean_zip(source_dir: Path, output_zip: Path):
    """
    创建干净的 zip 包

    @param source_dir: 要打包的目录
    @param output_zip: 输出 zip 文件路径
    """
    if output_zip.exists():
        output_zip.unlink()

    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        file_count = 0
        for path in source_dir.rglob('*'):
            if should_exclude(path):
                continue
            # 写入 zip：保留顶层目录名，去掉源路径前缀
            arcname = Path(FOLDER_NAME) / path.relative_to(source_dir)
            if path.is_file():
                zf.write(path, arcname)
                file_count += 1

    # 设置 zip 注释（Windows 友好）
    print(f'✅ 打包完成！共 {file_count} 个文件')
    print(f'📦 输出: {output_zip}')
    print(f'📏 大小: {output_zip.stat().st_size / 1024:.1f} KB')


def main():
    """主入口"""
    script_dir = Path(__file__).parent.resolve()
    source = script_dir
    output = script_dir.parent / f'{FOLDER_NAME}.zip'

    print('=' * 50)
    print('📦 打包工具（Windows 友好版）')
    print('=' * 50)
    print(f'源目录: {source}')
    print(f'输出:   {output}')
    print()

    create_clean_zip(source, output)

    print()
    print('💡 Windows 用户解压方式：')
    print('   1. 把 zip 发过去')
    print('   2. 右键 → 全部解压')
    print('   3. 双击里面的 【Windows系统请双击】运行对账.bat')


if __name__ == '__main__':
    main()
