import fitz  # PyMuPDF

def pdf_to_images(pdf_path, output_folder):
    pdf_document = fitz.open(pdf_path)

    for page_number in range(pdf_document.page_count):
        page = pdf_document.load_page(page_number)
        image = page.get_pixmap(dpi=300)  # 设置DPI为300，可以根据需要调整此值
        image_path = f"{output_folder}/page_{page_number+1}.png"  # 图片格式为PNG
        image.save(image_path)

    pdf_document.close()

if __name__ == "__main__":
    input_pdf_path = "C:/Users/7/PycharmProjects/ppt2png/Converted/线稿湖大.pdf"  # 替换为您的PDF文件路径
    output_folder_path = "C:/Users/7/PycharmProjects/ppt2png/Converted"  # 替换为您想要保存图像的文件夹路径

    pdf_to_images(input_pdf_path, output_folder_path)
