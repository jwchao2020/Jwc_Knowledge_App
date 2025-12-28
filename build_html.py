import os
import json
import subprocess
import shutil

# === ⚙️ 配置区域 ===
ROOT_DIR = os.getcwd()
SOURCE_DIR = os.path.join(ROOT_DIR, "source_word")     # 源文件目录
DOCS_DIR = os.path.join(ROOT_DIR, "Documents")         # App文档根目录
OUTPUT_DIR = os.path.join(DOCS_DIR, "content")         # HTML输出目录
MEDIA_DIR = os.path.join(DOCS_DIR, "media")            # 图片输出目录
DATA_JS_PATH = os.path.join(DOCS_DIR, "data.js")       # 目录数据文件

def init_folders():
    """初始化清理目录"""
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR)
    
    if os.path.exists(MEDIA_DIR):
        shutil.rmtree(MEDIA_DIR)
    os.makedirs(MEDIA_DIR)

def convert_files():
    tree_data = []
    
    # 遍历 source_word 文件夹
    for root, dirs, files in os.walk(SOURCE_DIR):
        # 排序，保证目录顺序
        dirs.sort()
        files.sort()
        
        for file in files:
            # 忽略临时文件
            if file.startswith("~"): continue

            src_path = os.path.join(root, file)
            # 计算相对路径，用于保持目录结构
            rel_path = os.path.relpath(src_path, SOURCE_DIR)
            rel_folder = os.path.dirname(rel_path)
            
            # 目标 HTML 文件夹
            target_folder = os.path.join(OUTPUT_DIR, rel_folder)
            if not os.path.exists(target_folder):
                os.makedirs(target_folder)

            file_name_no_ext = os.path.splitext(file)[0]
            output_html_path = os.path.join(target_folder, file_name_no_ext + ".html")
            
            # === 核心转换逻辑 ===
            cmd = []
            
            # 1. 处理 Word (.docx)
            if file.endswith(".docx"):
                print(f"转换 Word: {rel_path}")
                cmd = [
                    "pandoc", src_path,
                    "-f", "docx",
                    "-t", "html5",
                    "--mathjax",  # 处理公式
                    f"--extract-media={DOCS_DIR}", # 提取图片到 Documents/media
                    "-o", output_html_path
                ]
            
            # 2. 处理 Markdown (.md)
            elif file.endswith(".md"):
                print(f"转换 Markdown: {rel_path}")
                cmd = [
                    "pandoc", src_path,
                    "-f", "markdown",
                    "-t", "html5",
                    "--mathjax",
                    "-o", output_html_path
                ]
            
            else:
                continue # 跳过其他文件

            # 执行转换命令
            try:
                subprocess.run(cmd, check=True)
                
                # 添加到目录树
                # 注意：App读取时的路径是相对于 Documents/ 的
                web_path = os.path.join("content", rel_folder, file_name_no_ext + ".html")
                tree_data.append({
                    "title": file_name_no_ext,
                    "path": web_path,
                    "folder": rel_folder # 辅助字段，用于分组
                })
            except Exception as e:
                print(f"❌ 错误: {e}")

    return tree_data

def generate_js(data):
    # 生成 data.js
    # 这里做简单的扁平列表，如果需要多级折叠目录，需要更复杂的递归处理
    # 为了配合新的 index.html，我们把数据结构做成 { list: [...] }
    json_str = json.dumps(data, ensure_ascii=False, indent=2)
    content = f"const LOCAL_DATA = {{ version: '{os.times()}', list: {json_str} }};"
    
    with open(DATA_JS_PATH, 'w', encoding='utf-8') as f:
        f.write(content)
    print("✅ data.js 生成完成")

if __name__ == "__main__":
    print("🚀 开始构建 HTML 知识库...")
    init_folders()
    data = convert_files()
    generate_js(data)
    print("🎉 全部完成！")