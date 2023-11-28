import sys
import comtypes.client
import os
from PyPDF2 import PdfMerger
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel

class MyApp(QWidget):
    directoryname = None 
    def __init__(self):
        super().__init__()

        # 初始化UI界面
        self.initUI()

    def initUI(self):
        # 设置窗口标题和大小
        self.setWindowTitle('ppt-pdf转换器')
        self.setGeometry(300, 300, 300, 200)

        # 创建一个垂直布局
        layout = QVBoxLayout()

        # 创建按钮
        self.btn_selectDirectory = QPushButton('选择存放ppt的目录', self)
        self.btn_ppt2pdf = QPushButton('转换为pdf', self)
        self.btn_ppt2pdfandmerge = QPushButton('转换为pdf并自动合并', self)

        # 将按钮添加到布局中
        layout.addWidget(self.btn_selectDirectory)
        layout.addWidget(self.btn_ppt2pdf)
        layout.addWidget(self.btn_ppt2pdfandmerge)

        # 设置按钮的点击事件
        self.btn_selectDirectory.clicked.connect(self.selectDirectoryDialog)
        self.btn_ppt2pdf.clicked.connect(self.ppt2pdfUI)
        self.btn_ppt2pdfandmerge.clicked.connect(self.ppt2pdfandmergeUI)

        # 设置布局
        self.setLayout(layout)

    def selectDirectoryDialog(self):
        # 打开文件选择对话框
        options = QFileDialog.Options()
        dierctory = QFileDialog.getExistingDirectory(self, "选择存放ppt的目录", "", options=options)
        if dierctory:
            print(dierctory)
            self.directoryname = dierctory
    def ppt2pdf(self,filename,output_filename):
    #PPT文件导出为pdf格式
    #filename: PPT文件的名称
    #output_filename: 导出的pdf文件的名称

    # 2). 打开PPT程序
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1

    # 3). 通过PPT的应用程序打开指定的PPT文件
    # filename = "D:/大三下课程资料/嵌入式/PPT/xxx.ppt"
    # output_filename = "D:/大三下课程资料/嵌入式/PDF/xxx.pdf"
    
        presentation = powerpoint.Presentations.Open(filename)
        presentation.SaveAs(output_filename, 32)  # 32 表示将文件保存为PDF格式
        presentation.Close()
        print(output_filename+"导出成pdf格式成功!!!")
        powerpoint.Quit()

    def ppt2pdfUI(self):
        if(self.directoryname):
            # 确认按钮事件处理
            print('Confirmed')
            filenames = os.listdir(self.directoryname) #D:/大三下课程资料/嵌入式/PPT里面是xxx.ppt
            # for循环依次访问指定目录的所有文件名
            for filename in filenames:
                print(filename)
            # 判断文件的类型，对所有的ppt文件进行处理(ppt文件以ppt或者pptx结尾的)
                if filename.endswith('ppt') or filename.endswith('pptx'):
                    print(filename)           # PPT.pptx -> PPT.pdf
                    # 将filename以.进行分割，返回2个信息，文件的名称和文件的后缀名
                    base, ext = filename.split('.')  # base=PPT素材1 ext=pdf
                    new_name = base + '.pdf'         # PPT素材1.pdf
                    # ppt文件的完整位置:D:/大三下课程资料/嵌入式/PPT/xxx.ppt
                    filename = self.directoryname + '/' + filename
                    #print(filename)
                    # pdf文件的完整位置:D:/大三下课程资料/嵌入式/PDF/xxx.pdf
                    output_filename = self.directoryname + '/PDF/' + new_name
                    output_filename = output_filename.replace('/',"\\")
                    #print(output_filename)
                    # 将ppt转成pdf文件
                    self.ppt2pdf(filename, output_filename)
        else:
            print("还没有选择存放ppt的路径")
    def ppt2pdfandmergeUI(self):
        # 取消按钮事件处理
        print('Cancelled')
        if(self.directoryname):
            self.ppt2pdfUI()
            # 确认按钮事件处理
            print('Confirmed')
            newdirectoryname = self.directoryname + '/PDF'
            filenames = [filename for filename in os.listdir(newdirectoryname) if filename.endswith(".pdf")]  #D:/大三下课程资料/嵌入式/PPT里面是xxx.ppt
            # for循环依次访问指定目录的所有文件名
            
            #创建一个PdfMerger对象
            merger = PdfMerger()
 
            #逐个合并PDF文件
            for filename in filenames:
                print(filename)
                pdf_path = os.path.join(newdirectoryname, filename)
                merger.append(pdf_path)
 
            #指定合并后的PDF文件路径
            output_path = newdirectoryname+"/合并结果.pdf"
 
            merger.write(output_path)


        else:
            print("还没有选择存放ppt的路径")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    ex.show()
    sys.exit(app.exec_())