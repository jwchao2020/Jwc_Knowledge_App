import os
import json
import shutil
import time
import mammoth
import markdown

# === 配置 ===
SOURCE_DIR = "source_word"  # 你的源文件目录
OUTPUT_DIR = "Documents/content" # 转换后的 HTML 存放目录
DATA_FILE = "Documents/data.js"  # 索引文件

# 确保输出目录存在，如果存在则清空，防止旧文件干扰
if os.path.exists(OUTPUT_DIR):
    shutil.rmtree(OUTPUT_DIR)
os.makedirs(OUTPUT_DIR)

def convert_docx(src_path, dest_path):
    """转换 Docx -> HTML"""
    try:
        with open(src_path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
            # 简单的样式美化
            html = f"""
            <html><head>
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body {{ font-family: sans-serif; line-height: 1.6; padding: 15px; max-width: 800px; margin: 0 auto; }}
                img {{ max-width: 100%; height: auto; }}
                table {{ border-collapse: collapse; width: 100%; }}
                td, th {{ border: 1px solid #ddd; padding: 8px; }}
            </style>
            </head><body>{html}</body></html>
            """
            with open(dest_path, "w", encoding="utf-8") as f:
                f.write(html)
        return True
    except Exception as e:
        print(f"❌ 转换失败: {src_path} \n   原因: {e}")
        return False

def convert_md(src_path, dest_path):
    """转换 Markdown -> HTML"""
    try:
        with open(src_path, "r", encoding="utf-8") as f:
            text = f.read()
            html = markdown.markdown(text, extensions=['tables', 'fenced_code'])
            # 简单的样式美化
            html = f"""
            <html><head>
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body {{ font-family: sans-serif; line-height: 1.6; padding: 15px; color: #333; }}
                code {{ background: #f4f4f4; padding: 2px 5px; border-radius: 3px; }}
                pre {{ background: #f4f4f4; padding: 10px; overflow-x: auto; }}
                img {{ max-width: 100%; }}
                blockquote {{ border-left: 4px solid #ccc; margin: 0; padding-left: 10px; color: #666; }}
            </style>
            </head><body>{html}</body></html>
            """
            with open(dest_path, "w", encoding="utf-8") as f:
                f.write(html)
        return True
    except Exception as e:
        print(f"❌ 转换 Markdown 失败: {src_path} \n   原因: {e}")
        return False

def process_directory(current_src, current_dest, relative_root=""):
    """
    递归处理文件夹
    current_src: 当前源文件夹路径
    current_dest: 当前目标文件夹路径
    relative_root: 用于生成 URL 的相对路径
    """
    nodes = []
    
    # 获取当前目录下的所有条目，并排序（保证 0_, 1_ 顺序正确）
    try:
        items = sorted(os.listdir(current_src))
    except FileNotFoundError:
        return []

    for item in items:
        # 忽略隐藏文件
        if item.startswith('.'):
            continue

        src_path = os.path.join(current_src, item)
        dest_path = os.path.join(current_dest, item)
        
        # === 情况 1: 是文件夹 ===
        if os.path.isdir(src_path):
            # 在 content 下创建对应的文件夹
            if not os.path.exists(dest_path):
                os.makedirs(dest_path)
            
            # 递归处理子目录！
            children = process_directory(src_path, dest_path, os.path.join(relative_root, item))
            
            # 只有当文件夹里有内容时，才添加到目录树
            if children:
                nodes.append({
                    "title": item,  # 文件夹名字
                    "children": children # 子节点列表
                })
        
        # === 情况 2: 是文件 ===
        else:
            file_name, ext = os.path.splitext(item)
            ext = ext.lower()
            
            target_file_name = file_name + ".html"
            target_full_path = os.path.join(current_dest, target_file_name)
            web_path = "content/" + os.path.join(relative_root, target_file_name).replace("\\", "/")

            if ext == ".docx":
                print(f"📄 转换 Docx: {item}")
                if convert_docx(src_path, target_full_path):
                    nodes.append({
                        "title": file_name,
                        "path": web_path,
                        "type": "file"
                    })
            
            elif ext == ".md":
                print(f"📝 转换 MD: {item}")
                if convert_md(src_path, target_full_path):
                    nodes.append({
                        "title": file_name,
                        "path": web_path,
                        "type": "file"
                    })
            
            elif ext == ".pdf":
                # PDF 不转换，直接复制
                print(f"📑 复制 PDF: {item}")
                shutil.copy2(src_path, dest_path)
                # PDF 保持原名
                web_path_pdf = "content/" + os.path.join(relative_root, item).replace("\\", "/")
                nodes.append({
                    "title": file_name,
                    "path": web_path_pdf,
                    "type": "pdf"
                })

    return nodes

# === 主程序 ===
print("🚀 开始构建目录树...")
tree_structure = process_directory(SOURCE_DIR, OUTPUT_DIR)

# 生成 JSON
data = {
    "version": int(time.time()),
    "tree": tree_structure
}

with open(DATA_FILE, "w", encoding="utf-8") as f:
    f.write(f"const LOCAL_DATA = {json.dumps(data, ensure_ascii=False, indent=2)};")

print(f"✅ 构建完成！索引已保存至 {DATA_FILE}")