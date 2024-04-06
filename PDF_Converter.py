import os
import win32com.client

class pdfConverter:
    def __init__(self,folder):
        self.pptFormatPDF = 32
        self.pptToPDF = win32com.client.Dispatch("PowerPoint.Application")
        self.pptToPDF.Visible = 1 
        self.pdf_Path = os.path.join(folder, "./Converted")
        self.format_document()

    def ppt_to_pdf(self):
        files = os.listdir(folder)
        ppt_files = [f for f in files if f.endswith((".ppt",".pptx"))]
        print("共有 {} 个文件待转换。".format(len(ppt_files)))
        for index, ppt_file in enumerate(ppt_files):
            print("正在转换第{}个文件: {}".format(index+1, ppt_file))
            ppt_Path = os.path.join(folder, ppt_file)
            if '.pptx' in ppt_file:
                ppt_file = ppt_file.replace('pptx', 'pdf')
            else:
                ppt_file = ppt_file.replace('ppt', 'pdf')
            pdf_Path = os.path.join(self.pdf_Path,ppt_file)
            pdfCreate = self.pptToPDF.Presentations.Open(ppt_Path)
            pdfCreate.SaveAs(pdf_Path, self.pptFormatPDF)
            pdfCreate.Close()
        self.pptToPDF.Quit()

            
    def format_document(self):
        isExist = os.path.exists(self.pdf_Path)
        if not isExist:
            os.makedirs(self.pdf_Path)

folder = os.getcwd()
Converter = pdfConverter(folder)
Converter.ppt_to_pdf()
os._exit(1)
