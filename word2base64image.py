import os
import subprocess
import fitz  # PyMuPDF
from PIL import Image
import base64
from io import BytesIO


def word_to_base64image(word_path: str, zoom_x: float = 2.0, zoom_y: float = 2.0) -> tuple:
    """
    将 Word 文档转换为 PDF → 再转为垂直拼接图片 → 生成 base64，
    所有文件保存在与 Word 同名的子文件夹中。

    参数:
        word_path (str): Word 文件路径（如 "documents/demo.docx"）
        zoom_x, zoom_y (float): PDF 渲染缩放倍率（默认 2.0 ≈ 300 DPI）

    返回:
        tuple: (pdf_path, image_path, base64_txt_path)

    生成结构示例：
        documents/demo/
            demo.pdf
            img.jpg
            base64.txt
    """
    if not os.path.exists(word_path):
        raise FileNotFoundError(f"Word 文件未找到: {word_path}")

    # 获取路径信息
    word_dir = os.path.dirname(word_path)
    word_name = os.path.basename(word_path)
    name_without_ext = os.path.splitext(word_name)[0]

    # 创建输出子文件夹（以 Word 文件名命名）
    output_folder = os.path.join(word_dir, name_without_ext)
    os.makedirs(output_folder, exist_ok=True)

    # 定义输出文件路径
    pdf_path = os.path.join(output_folder, f"{name_without_ext}.pdf")
    image_path = os.path.join(output_folder, "img.jpg")
    base64_txt_path = os.path.join(output_folder, "base64.txt")

    # ==================== 第一步：Word → PDF ====================
    print(f"正在将 Word 转换为 PDF: {word_path} → {pdf_path}")

    # 方法1：LibreOffice（推荐）
    converted = False
    try:
        result = subprocess.run([
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_folder,
            word_path
        ], check=True, capture_output=True, timeout=120)

        # LibreOffice 生成的 PDF 文件名与原 Word 一致
        generated_pdf = os.path.join(output_folder, name_without_ext + ".pdf")
        if os.path.exists(generated_pdf):
            if generated_pdf != pdf_path:
                os.replace(generated_pdf, pdf_path)
            print("✓ Word → PDF 转换成功（LibreOffice）")
            converted = True

    except FileNotFoundError:
        print("LibreOffice 未安装或不可用，尝试 Microsoft Word 方式...")
    except subprocess.TimeoutExpired:
        print("LibreOffice 转换超时")
    except subprocess.CalledProcessError as e:
        print(f"LibreOffice 转换失败: {e.stderr.decode('utf-8', errors='ignore').strip()}")

    # 方法2：Windows 下使用 Microsoft Word COM
    if not converted and os.name == 'nt':
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            wdFormatPDF = 17

            doc = word.Documents.Open(os.path.abspath(word_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=wdFormatPDF)
            doc.Close()
            word.Quit()

            print("✓ Word → PDF 转换成功（Microsoft Word）")
            converted = True
        except ImportError:
            print("pywin32 未安装，无法使用 Word COM")
        except Exception as e:
            print(f"Word COM 转换失败: {e}")

    if not converted:
        raise RuntimeError("所有 Word → PDF 转换方式均失败，请检查 LibreOffice 或 Microsoft Office 安装")

    # ==================== 第二步：PDF → 拼接图片 + base64 ====================
    print(f"正在将 PDF 转换为图片并生成 base64: {pdf_path}")

    pdf = fitz.open(pdf_path)
    images = []

    for page_num in range(len(pdf)):
        page = pdf.load_page(page_num)
        mat = fitz.Matrix(zoom_x, zoom_y)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        images.append(img)

    pdf.close()

    if not images:
        raise ValueError("PDF 中没有可转换的页面")

    # 垂直拼接
    total_height = sum(img.height for img in images)
    max_width = max(img.width for img in images)
    concatenated = Image.new('RGB', (max_width, total_height), (255, 255, 255))

    y_offset = 0
    for img in images:
        x_offset = (max_width - img.width) // 2
        concatenated.paste(img, (x_offset, y_offset))
        y_offset += img.height

    # 保存图片
    concatenated.save(image_path, 'JPEG', quality=95)
    print(f"✓ 拼接图片已保存: {image_path}")

    # 生成 base64
    buffered = BytesIO()
    concatenated.save(buffered, format="JPEG", quality=95)
    base64_str = base64.b64encode(buffered.getvalue()).decode('utf-8')
    data_url = f"data:image/jpeg;base64,{base64_str}"

    with open(base64_txt_path, 'w', encoding='utf-8') as f:
        f.write(data_url)

    print(f"✓ Base64 已保存: {base64_txt_path}")

    return pdf_path, image_path, base64_txt_path


# ==================== 示例用法 ====================
if __name__ == "__main__":
    # 批量处理多个 Word 文件
    word_files = [
        "documents/demo.docx"
        # 多个文件路径
    ]

    for word_file in word_files:
        try:
            pdf_p, img_p, b64_p = word_to_base64image(word_file)
            print(f"完成: {word_file} → {os.path.dirname(pdf_p)}/\n")
        except Exception as e:
            print(f"处理失败 {word_file}: {e}\n")