import os
import comtypes.client
import getlist

def ppt_to_images(input_ppt_file, output_folder):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(input_ppt_file)
    slides_count = len(presentation.Slides)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for i in range(1, slides_count + 1):
        slide = presentation.Slides(i)
        slide.Export(os.path.join(output_folder, f"slide_{i}.png"), "PNG")

    presentation.Close()
    powerpoint.Quit()

if __name__ == "__main__":
    directory_path = r"D:\PPT2PNG\PPT\PPT模板"  # 指定要获取文件名的目录路径

    all_filenames = getlist.get_all_filenames_in_directory(directory_path)
    for ppt in all_filenames:

        input_ppt_file = f"D:/PPT2PNG/PPT/PPT模板/{ppt}"  # 输入PPT文件的路径
        try:
            os.makedirs(f"D:/PPT2PNG/PPT/{ppt}"[:-5])
        except FileExistsError:
            pass
        output_folder = f"D:/PPT2PNG/PPT/{ppt}"[:-5]  # 输出图片保存的文件夹路径

        ppt_to_images(input_ppt_file, output_folder)
