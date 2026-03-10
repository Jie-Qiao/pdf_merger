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

class DragDropListWidget(QListWidget):
    """支持拖拽的列表组件"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    file_path = str(url.toLocalFile())
                    # 避免重复添加
                    items = [self.item(i).text() for i in range(self.count())]
                    if file_path not in items:
                        self.addItem(file_path)
        else:
            event.ignore()

class PDFMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.temp_dir = tempfile.gettempdir()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('全能文档转PDF合并器')
        self.resize(550, 450)
        layout = QVBoxLayout()

        # 提示标签
        self.label = QLabel("请将 Word, Excel, 图片 或 OFD 文件拖入下方列表中\n(拖动可排序，按 Delete 键删除)")
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
                # 遍历所有工作表，强行设置打印缩放比例
                for sheet in wb.Worksheets:
                    try:
                        # 关闭固定缩放比例
                        sheet.PageSetup.Zoom = False
                        # 强行将所有列宽压缩到 1 页宽内
                        sheet.PageSetup.FitToPagesWide = 1
                        # 高度不限制（False），让行数自然往下排，避免内容被压扁
                        sheet.PageSetup.FitToPagesTall = False 
                    except Exception:
                        # 如果工作表被密码保护，可能会设置失败，这里直接跳过保护表
                        pass
                # -------------------------------------------
                
                wb.ExportAsFixedFormat(0, os.path.abspath(output_path))
                
                # SaveChanges=False 非常重要！因为修改了 PageSetup，不加这个关文件时会卡住弹窗问你要不要保存
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