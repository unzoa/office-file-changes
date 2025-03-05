import os
import subprocess
from pathlib import Path

# 添加基础目录计算（在函数外部）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def convert_with_libreoffice(src_path, dest_dir):
    """使用LibreOffice进行文档转换"""
    try:
        cmd = [
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', dest_dir,
            src_path
        ]
        subprocess.run(cmd, check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"转换失败: {e}")
        return False

def convert_word_to_pdf(source_dir, output_dir):
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    for root, _, files in os.walk(source_dir):
        for file in files:
            if file.lower().endswith(('.doc', '.docx')):
                src_path = os.path.join(root, file)
                dest_dir = os.path.join(output_dir, os.path.relpath(root, source_dir))

                Path(dest_dir).mkdir(parents=True, exist_ok=True)

                if convert_with_libreoffice(src_path, dest_dir):
                    pdf_name = os.path.splitext(file)[0] + '.pdf'
                    print(f"转换成功: {pdf_name}")

if __name__ == "__main__":
    # 使用计算后的基础路径
    source = os.path.join(BASE_DIR, "docs")
    output = os.path.join(BASE_DIR, "pdfs")

    if not os.path.exists(source):
        print(f"源目录不存在：{source}")
        exit(1)

    convert_word_to_pdf(source, output)