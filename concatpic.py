from PIL import Image

def combine_images(image_paths, output_path, rows, cols):
    images = [Image.open(path) for path in image_paths]

    # 组合图片的尺寸
    combined_width = max(img.width for img in images) * cols
    combined_height = max(img.height for img in images) * rows

    # 创建一个新的空白图片
    combined_image = Image.new('RGB', (combined_width, combined_height))

    current_row, current_col = 0, 0

    for img in images:
        combined_image.paste(img, (current_col * img.width, current_row * img.height))

        current_col += 1
        if current_col == cols:
            current_col = 0
            current_row += 1

    combined_image.save(output_path)
    # print(f"Images combined and saved to {output_path}")

def combine_images2(image_path1, image_path2, output_path):
    image1 = Image.open(image_path1)
    image2 = Image.open(image_path2)

    # 调整图片大小
    image1 = image1.resize((image1.width * 2, image1.height * 2))

    # 创建一个新的空白图片
    combined_image = Image.new('RGB', (image2.width, image2.height + image1.height))

    combined_image.paste(image1, (0, 0))
    combined_image.paste(image2, (0, image1.height))

    combined_image.save(output_path)
    print(f"Images combined and saved to {output_path}")


if __name__ == "__main__":
    folder_name = 'Ticket Sales App Pitch Deck by Slidesgo'
    image_paths = [f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片2.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片3.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片4.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片5.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片6.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片7.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片8.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片9.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片10.JPG",
                   f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片11.JPG"
                   ]  # 输入要组合的图片路径列表
    output_path1 = f"C:/Users/7/Desktop/ppt2png/{folder_name}/combined_image.jpg"  # 输出组合后的图片路径
    output_path2 = f"C:/Users/7/Desktop/ppt2png/{folder_name}/combined_image.jpg"
    rows = 5  # 行数
    cols = 2  # 列数

    image_path1 = f"C:/Users/7/Desktop/ppt2png/{folder_name}/幻灯片1.JPG"
    image_path2 = f"C:/Users/7/Desktop/ppt2png/{folder_name}/combined_image.jpg"

    combine_images(image_paths, output_path1, rows, cols)
    combine_images2(image_path1, image_path2, output_path2)
