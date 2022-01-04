# coding: utf-8
import sys
from PyQt5.QtWidgets import *
from myUI import Ui_MainWindow

import docx
import pathlib

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.words_list = []
        self.replace_text_list = {}
        self.ui.addWordBtn.clicked.connect(self.addWordBtn_clicked)
        self.ui.replaceBtn.clicked.connect(self.replaceBtn_clicked)
        self.ui.replaceTextAddBtn.clicked.connect(self.replaceTextAddBtn_clicked)
        self.ui.replaceText.returnPressed.connect(self.ui.replaceTextAddBtn.click)

    def addWordBtn_clicked(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        files, _ = QFileDialog.getOpenFileNames(self,"QFileDialog.getOpenFileNames()", "","Word文件 (*.docx)", options=options)
        if files:
            for file in files:
                if file not in self.words_list:
                    self.words_list.append(file)
                    self.ui.wordList.addItem(file)
                    
    def replaceBtn_clicked(self):
        for file_full_path in self.words_list:
            path = pathlib.Path(file_full_path)
            file_name = path.name
            doc = docx.Document(file_full_path)
            
            for para in doc.paragraphs:
                for item in self.replace_text_list.items():
                    search_text = item[0]
                    replace_text = item[1]
                    if search_text in para.text:
                        inline = para.runs
                        for i in range(len(inline)):
                            if search_text in inline[i].text:
                                text = inline[i].text.replace(search_text, replace_text)
                                inline[i].text = text
            doc.save(f'output/新 {file_name}')
        self.ui.progressBar.setValue(100)
        
    def replaceTextAddBtn_clicked(self):
        search_text = self.ui.searchText.text()
        replace_text = self.ui.replaceText.text()
        
        self.ui.searchText.setText("")
        self.ui.replaceText.setText("")
        
        if search_text != replace_text:
            self.ui.replaceList.addItem(f"{search_text} => {replace_text}")
            self.replace_text_list[search_text] = replace_text

if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())