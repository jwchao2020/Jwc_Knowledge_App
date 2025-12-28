import os
import json
import base64
import shutil
import mammoth  # 专门用于 docx 转 html
import markdown # 专门用于 md 转 html

# === ⚙️ 配置区域 ===
ROOT_DIR = os.getcwd()
SOURCE_DIR = os.path.join(ROOT_DIR, "source_word")
DOCS_DIR = os.path.join(ROOT_DIR, "Documents")
OUTPUT_DIR = os.path.join(DOCS_DIR, "content")
MEDIA_DIR = os.path.join(DOCS_DIR, "media")
DATA_JS_PATH = os.path.join(DOCS_DIR, "data.js")

def init_folders():
    if os.path.exists(OUTPUT_DIR): shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR)
    # mammoth 会直接把图片转为 base64 内嵌在 html 里，所以 media 目录其实不强制需要，但保留结构
    if not os.path.exists(MEDIA_DIR): os.makedirs(MEDIA_DIR)

def convert_image(image):
    # Mammoth 图片处理：转为 Base64 内嵌，防止路径丢失问题
    with image.open() as image_bytes:
        encoded_src = base64.b64encode(image_bytes.read()).decode("ascii")
    return {
        "src": "data:" + image.content_type + ";base64," + encoded_src
    }

def convert_files():
    tree_data = []
    
    for root, dirs, files in os.walk(SOURCE_DIR):
        dirs.sort()
        files.sort()
        
        for file in files:
            if file.startswith("~") or file.startswith("."): continue

            src_path = os.path.join(root, file)
            rel_path = os.path.relpath(src_path, SOURCE_DIR)
            rel_folder = os.path.dirname(rel_path)
            target_folder = os.path.join(OUTPUT_DIR, rel_folder)
            
            if not os.path.exists(target_folder): os.makedirs(target_folder)

            file_name_no_ext = os.path.splitext(file)[0]
            output_html_path = os.path.join(target_folder, file_name_no_ext + ".html")
            
            print(f"正在转换: {rel_path} ...", end="")

            try:
                html_content = ""
                
                # === 方案 A: Word 转 HTML (使用 Mammoth) ===
                if file.endswith(".docx"):
                    with open(src_path, "rb") as docx_file:
                        # style_map 自定义样式，让表格和图片更好看
                        style_map = """
                        p[style-name='Heading 1'] => h1:fresh
                        p[style-name='Heading 2'] => h2:fresh
                        p[style-name='Heading 3'] => h3:fresh
                        table => table.table.table-bordered
                        """
                        result = mammoth.convert_to_html(
                            docx_file, 
                            convert_image=mammoth.images.img_element(convert_image),
                            style_map=style_map
                        )
                        html_content = result.value
                        messages = result.messages # 警告信息

                # === 方案 B: Markdown 转 HTML ===
                elif file.endswith(".md"):
                    with open(src_path, "r", encoding="utf-8") as md_file:
                        text = md_file.read()
                        html_content = markdown.markdown(text, extensions=['tables', 'fenced_code'])

                else:
                    print(" [跳过]")
                    continue

                # 写入 HTML 文件
                # 额外包裹一层 div 以便 CSS 样式生效
                final_html = f'<div class="doc-container">{html_content}</div>'
                
                with open(output_html_path, "w", encoding="utf-8") as f:
                    f.write(final_html)

                # 添加到目录索引
                web_path = os.path.join("content", rel_folder, file_name_no_ext + ".html")
                tree_data.append({
                    "title": file_name_no_ext,
                    "path": web_path,
                    "folder": rel_folder if rel_folder else "ROOT"
                })
                print(" ✅ 成功")

            except Exception as e:
                print(f"\n❌ 失败! 文件可能已损坏: {src_path}")
                print(f"   错误信息: {e}")
                # 即使失败，也生成一个报错的 HTML，方便在 App 里看到哪个文件坏了
                error_html = f'<h3 style="color:red">文件转换失败</h3><p>该文档可能已损坏或格式不兼容。</p><pre>{str(e)}</pre>'
                with open(output_html_path, "w", encoding="utf-8") as f:
                    f.write(error_html)
                
                # 依然添加到目录，这样你在 App 里能看到它
                web_path = os.path.join("content", rel_folder, file_name_no_ext + ".html")
                tree_data.append({
                    "title": f"⚠️ {file_name_no_ext} (损坏)",
                    "path": web_path,
                    "folder": rel_folder
                })

    return tree_data

def generate_js(data):
    json_str = json.dumps(data, ensure_ascii=False, indent=2)
    # 使用 os.times() 可能在不同系统不一致，改用时间戳字符串
    import time
    ver = str(int(time.time()))
    content = f"const LOCAL_DATA = {{ version: '{ver}', list: {json_str} }};"
    
    with open(DATA_JS_PATH, 'w', encoding='utf-8') as f:
        f.write(content)
    print("✅ data.js 索引已更新")

if __name__ == "__main__":
    print("🚀 使用 Python Native 模式构建知识库...")
    init_folders()
    data = convert_files()
    generate_js(data)