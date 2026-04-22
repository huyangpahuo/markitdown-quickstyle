import os
import sys
import shutil
import zipfile
import tempfile
from pathlib import Path
from markitdown import MarkItDown

# 支持的图片格式
IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.gif', '.webp', '.svg', '.bmp', '.tiff'}

def inject_yaml_and_save(md_content, out_md_path, assets_folder_name):
    """注入 Typora YAML 头并保存"""
    yaml_header = f"---\ntypora-copy-images-to: ./{assets_folder_name}\n---\n\n"
    with open(out_md_path, 'w', encoding='utf-8') as f:
        f.write(yaml_header + md_content)

def extract_office_images(file_path, assets_dir):
    """暴力提取 Office (docx/pptx/xlsx) 中的媒体图片到 assets 文件夹"""
    ext = Path(file_path).suffix.lower()
    if ext not in ['.docx', '.pptx', '.xlsx']:
        return
    
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            for info in z.infolist():
                # 匹配 word/media/, ppt/media/, xl/media/ 目录下的文件
                if '/media/' in info.filename and not info.filename.endswith('/'):
                    extracted_path = z.extract(info, path=tempfile.gettempdir())
                    shutil.copy(extracted_path, assets_dir)
    except Exception as e:
        print(f"  [警告] 提取 {ext} 内部图片时失败: {e}")

def convert_single_file(file_path, base_out_dir):
    """处理单个文件：用同名文件夹包裹 MD 和 assets"""
    file_path = Path(file_path)
    if not file_path.exists() or file_path.suffix.lower() == '.zip':
        return
    
    base_name = file_path.stem
    # 创建独立包裹文件夹: output/文件名/
    wrapper_dir = Path(base_out_dir) / base_name
    wrapper_dir.mkdir(parents=True, exist_ok=True)
    
    # 创建 assets 文件夹: output/文件名/文件名_assets/
    assets_folder_name = f"{base_name}_assets"
    assets_dir = wrapper_dir / assets_folder_name
    assets_dir.mkdir(parents=True, exist_ok=True)
    
    # 提取 Office 内部图片
    extract_office_images(file_path, assets_dir)
    
    print(f"  正在转换: {file_path.name} ...")
    try:
        md = MarkItDown()
        result = md.convert(str(file_path))
        md_content = result.text_content if result else ""
    except Exception as e:
        md_content = f"转换失败: {e}"
        
    out_md_path = wrapper_dir / f"{base_name}.md"
    inject_yaml_and_save(md_content, out_md_path, assets_folder_name)

def convert_merged_images(image_paths, base_out_dir, folder_name="新建文件"):
    """将多个图片合并为一个 MD 文件"""
    if not image_paths:
        return
        
    wrapper_dir = Path(base_out_dir) / folder_name
    wrapper_dir.mkdir(parents=True, exist_ok=True)
    
    assets_folder_name = f"{folder_name}_assets"
    assets_dir = wrapper_dir / assets_folder_name
    assets_dir.mkdir(parents=True, exist_ok=True)
    
    md_lines = []
    for img in image_paths:
        img = Path(img)
        # 复制图片到 assets
        shutil.copy(img, assets_dir)
        # 写入 MD 语法
        md_lines.append(f"![{img.stem}](./{assets_folder_name}/{img.name})\n")
        
    out_md_path = wrapper_dir / f"{folder_name}.md"
    inject_yaml_and_save("\n".join(md_lines), out_md_path, assets_folder_name)

def process_zip_depth1(zip_path, base_out_dir):
    """只处理 ZIP 根目录深度 1 的文件"""
    zip_path = Path(zip_path)
    zip_wrapper_dir = Path(base_out_dir) / zip_path.stem
    zip_wrapper_dir.mkdir(parents=True, exist_ok=True)
    
    temp_dir = Path(tempfile.mkdtemp())
    images_to_merge = []
    
    with zipfile.ZipFile(zip_path, 'r') as z:
        for info in z.infolist():
            # 核心规则：跳过文件夹，并且只处理根目录文件（文件名中没有 '/'）
            if info.is_dir() or '/' in info.filename:
                continue
                
            extracted_path = z.extract(info, path=temp_dir)
            ext = Path(extracted_path).suffix.lower()
            
            if ext in IMAGE_EXTS:
                images_to_merge.append(extracted_path)
            else:
                # 正常转换非图片文件，放入 ZIP 专属的包裹目录中
                convert_single_file(extracted_path, zip_wrapper_dir)
                
    # 合并处理压缩包根目录的所有图片
    if images_to_merge:
        convert_merged_images(images_to_merge, zip_wrapper_dir, "新建文件")
        
    shutil.rmtree(temp_dir, ignore_errors=True)

def main():
    if len(sys.argv) < 3:
        return
        
    mode = sys.argv[1]
    target = sys.argv[2]
    out_dir = sys.argv[3] if len(sys.argv) > 3 else "output"
    
    if mode == "single":
        if target.lower().endswith('.zip'):
            process_zip_depth1(target, out_dir)
        else:
            convert_single_file(target, out_dir)
            
    elif mode == "batch_images":
        # 批量处理 input 里的所有图片
        input_dir = Path(target)
        images = [f for f in input_dir.iterdir() if f.suffix.lower() in IMAGE_EXTS and f.is_file()]
        if images:
            print(f"  发现 {len(images)} 张图片，正在合并为 [新建文件.md] ...")
            convert_merged_images(images, out_dir, "新建文件")
            
    elif mode == "batch":
        # 批量处理 input 里的特定格式
        input_dir = Path(target)
        category_exts = sys.argv[4].split(',') if len(sys.argv) > 4 else []
        
        for f in input_dir.iterdir():
            if f.is_file() and f.suffix.lower() in category_exts:
                if f.suffix.lower() == '.zip':
                    process_zip_depth1(f, out_dir)
                else:
                    convert_single_file(f, out_dir)

if __name__ == "__main__":
    main()