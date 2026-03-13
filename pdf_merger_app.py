import sys
import os
import tempfile
import base64
import win32com.client
from PIL import Image
from pypdf import PdfWriter
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QListWidget, QLabel, QMessageBox, QAbstractItemView, QLineEdit, QFileDialog, QCheckBox)
from PyQt5.QtCore import Qt

# 根据官方示例引入 OFD
try:
    from easyofd.ofd import OFD
    HAS_EASYOFD = True
except ImportError:
    HAS_EASYOFD = False

# 全局定义支持的扩展名，方便拖拽文件夹时进行过滤
SUPPORTED_EXTS = ['.docx', '.doc', '.xlsx', '.xls', '.ppt', '.pptx', 
                  '.jpg', '.jpeg', '.png', '.bmp', '.ofd', '.pdf']

class DragDropListWidget(QListWidget):
    """支持拖拽的列表组件"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        # 修改为 DragDrop 模式，同时支持外部拖入和内部移动
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction) # 内部默认动作为移动
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    def dragEnterEvent(self, event):
        # 1. 如果是内部拖拽排序，交回给系统默认逻辑处理
        if event.source() == self:
            super().dragEnterEvent(event)
        # 2. 如果是从外部拖入文件/文件夹
        elif event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.source() == self:
            super().dragMoveEvent(event)
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        # 1. 处理内部拖拽排序
        if event.source() == self:
            super().dropEvent(event)
        # 2. 处理外部文件拖入
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = str(url.toLocalFile())
                    
                    # 文件夹递归处理
                    if os.path.isdir(path):
                        for root, dirs, files in os.walk(path):
                            for file in files:
                                ext = os.path.splitext(file)[1].lower()
                                if ext in SUPPORTED_EXTS:
                                    full_path = os.path.join(root, file)
                                    self._add_item_if_unique(full_path)
                    # 单文件处理
                    else:
                        ext = os.path.splitext(path)[1].lower()
                        if ext in SUPPORTED_EXTS:
                            self._add_item_if_unique(path)
        else:
            event.ignore()

    def _add_item_if_unique(self, file_path):
        """避免重复添加文件的辅助方法"""
        items = [self.item(i).text() for i in range(self.count())]
        if file_path not in items:
            self.addItem(file_path)

class PDFMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.temp_dir = tempfile.gettempdir()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('全能文档转PDF合并器')
        self.resize(550, 450)
        layout = QVBoxLayout()

        # 提示标签 (更新了文案，加入了 PPT 和文件夹提示)
        self.label = QLabel("请将 Word, Excel, PPT, 图片 或 OFD 的 文件/文件夹 拖入下方列表中\n(拖动可排序，按 Delete 键删除)")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        # 拖拽列表
        self.file_list = DragDropListWidget(self)
        layout.addWidget(self.file_list)

        # 清空按钮
        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.clicked.connect(self.file_list.clear)
        layout.addWidget(self.btn_clear)

        # --- 输出设置区域 ---
        settings_layout = QVBoxLayout()
        
        # 1. 目录选择
        dir_layout = QHBoxLayout()
        self.dir_label = QLabel("输出目录:")
        self.dir_input = QLineEdit(os.path.join(os.path.expanduser("~"), "Desktop")) # 默认桌面
        self.dir_btn = QPushButton("选择...")
        self.dir_btn.clicked.connect(self.select_directory)
        dir_layout.addWidget(self.dir_label)
        dir_layout.addWidget(self.dir_input)
        dir_layout.addWidget(self.dir_btn)
        settings_layout.addLayout(dir_layout)
        
        # 2. 预览选项
        self.cb_open = QCheckBox("合并完成后自动打开 PDF 进行打印预览")
        self.cb_open.setChecked(True) # 默认勾选
        settings_layout.addWidget(self.cb_open)
        
        layout.addLayout(settings_layout)

        # 合并按钮
        self.btn_generate = QPushButton("一键转换并合并")
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; height: 40px;")
        self.btn_generate.clicked.connect(self.process_files)
        layout.addWidget(self.btn_generate)

        self.setLayout(layout)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            for item in self.file_list.selectedItems():
                self.file_list.takeItem(self.file_list.row(item))

    def select_directory(self):
        """选择输出目录"""
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录", self.dir_input.text())
        if folder:
            self.dir_input.setText(folder)

    def process_files(self):
        count = self.file_list.count()
        if count == 0:
            QMessageBox.warning(self, "警告", "请先拖入需要处理的文件！")
            return

        out_dir = self.dir_input.text()
        if not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法创建输出目录:\n{str(e)}")
                return

        self.btn_generate.setText("正在处理中，请稍候...")
        self.btn_generate.setEnabled(False)
        QApplication.processEvents()

        pdf_files_to_merge = []
        merger = PdfWriter()
        output_pdf_path = os.path.join(out_dir, "合并输出_打印预览.pdf")

        try:
            for i in range(count):
                file_path = self.file_list.item(i).text()
                ext = os.path.splitext(file_path)[1].lower()
                temp_pdf = os.path.join(self.temp_dir, f"temp_{i}.pdf")
                
                if ext in ['.docx', '.doc']:
                    self.convert_word(file_path, temp_pdf)
                elif ext in ['.xlsx', '.xls']:
                    self.convert_excel(file_path, temp_pdf)
                elif ext in ['.ppt', '.pptx']: # 新增 PPT 分支
                    self.convert_powerpoint(file_path, temp_pdf)
                elif ext in ['.jpg', '.jpeg', '.png', '.bmp']:
                    self.convert_image(file_path, temp_pdf)
                elif ext == '.ofd':
                    self.convert_ofd(file_path, temp_pdf)
                elif ext == '.pdf':
                    temp_pdf = file_path
                else:
                    raise Exception(f"不支持的文件格式: {ext}")

                pdf_files_to_merge.append(temp_pdf)

            # 合并 PDF
            for pdf in pdf_files_to_merge:
                merger.append(pdf)
            merger.write(output_pdf_path)
            merger.close()

            self.file_list.clear()

            # 根据勾选状态决定是否打开
            if self.cb_open.isChecked():
                os.startfile(output_pdf_path)
            else:
                QMessageBox.information(self, "成功", f"文件已成功合并并保存至:\n{output_pdf_path}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理过程中出错:\n{str(e)}")
        finally:
            self.btn_generate.setText("一键转换并合并")
            self.btn_generate.setEnabled(True)

    # --- 各种格式转换核心逻辑 ---
    
    def convert_powerpoint(self, input_path, output_path):
        """新增：PowerPoint转PDF逻辑"""
        app = None
        presentation = None
        try:
            # 1. 优先尝试启动 Microsoft PowerPoint
            try:
                app = win32com.client.DispatchEx("PowerPoint.Application")
            except:
                # 2. 尝试 WPS 演示
                try:
                    app = win32com.client.DispatchEx("kwpp.Application")
                except:
                    app = win32com.client.DispatchEx("WPP.Application")
            
            # 打开 PPT：ReadOnly=1 (True), Untitled=0 (False), WithWindow=0 (False)
            # 以免弹出不需要的前台窗口
            presentation = app.Presentations.Open(os.path.abspath(input_path), 1, 0, 0)
            
            # 32 是导出为 PDF 的标准代码 (ppSaveAsPDF)
            presentation.SaveAs(os.path.abspath(output_path), 32)
        finally:
            if presentation:
                presentation.Close()
            if app:
                app.Quit()

    def convert_word(self, input_path, output_path):
        app = None
        try:
            # 1. 优先尝试启动 Microsoft Word
            try:
                app = win32com.client.DispatchEx("Word.Application")
            except:
                # 2. 如果失败，尝试启动 WPS 文字
                app = win32com.client.DispatchEx("kwps.Application")
            
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(os.path.abspath(input_path), ReadOnly=True)
            # 17 是导出为 PDF 的标准代码，WPS 也通用
            doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
            doc.Close()
        finally:
            if app:
                app.Quit()

    def convert_excel(self, input_path, output_path):
        app = None
        try:
            # 1. 优先尝试启动 Microsoft Excel
            try:
                app = win32com.client.DispatchEx("Excel.Application")
            except:
                # 2. 如果失败，尝试启动 WPS 表格 (优先 ket，兼容老版 ET)
                try:
                    app = win32com.client.DispatchEx("ket.Application")
                except:
                    app = win32com.client.DispatchEx("ET.Application")
            
            app.Visible = False
            app.DisplayAlerts = False 
            wb = app.Workbooks.Open(os.path.abspath(input_path), ReadOnly=True)
            
            # --- 核心修复：解决 Excel 打印冗余页问题 ---
            for sheet in wb.Worksheets:
                try:
                    sheet.PageSetup.Zoom = False
                    sheet.PageSetup.FitToPagesWide = 1
                    sheet.PageSetup.FitToPagesTall = False 
                except Exception:
                    pass
            # -------------------------------------------
            
            wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
            
            # SaveChanges=False 非常重要！
            wb.Close(SaveChanges=False) 
        finally:
            if app:
                app.Quit()

    def convert_image(self, input_path, output_path):
        image = Image.open(input_path)
        if image.mode != 'RGB':
            image = image.convert('RGB')
        image.save(output_path)
        
    def convert_ofd(self, input_path, output_path):
        """完全按照用户提供的官方示例重写的 OFD 转换逻辑"""
        if not HAS_EASYOFD:
            raise Exception("未安装 easyofd 库！请检查环境。")
            
        with open(input_path, "rb") as f:
            ofdb64 = str(base64.b64encode(f.read()), "utf-8")
            
        ofd = OFD()  # 初始化OFD 工具类
        ofd.read(ofdb64)  # 读取ofdb64，此处不生成多余的xml
        pdf_bytes = ofd.to_pdf()  # 转pdf
        ofd.del_data() # 清理内存

        with open(output_path, "wb") as f:
            f.write(pdf_bytes)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion') 
    ex = PDFMergerApp()
    ex.show()
    sys.exit(app.exec_())