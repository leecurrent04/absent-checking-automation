import pandas
import time
import sys
import os
import os.path
import shutil

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5 import uic
from PyQt5.QtCore import Qt


# Connect UI file
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, "resource/", relative_path)


form_main = uic.loadUiType(resource_path("main.ui"))[0]

path_csv = ""
path_directory = ""

# path_chrome = "google-chrome"
path_chrome = "C:\Program` Files\Google\Chrome\Application\chrome.exe"
path_msedge = "C:\Program` Files` `(x86`)\Microsoft\Edge\Application\msedge.exe"


class WindowMainClass(QMainWindow, form_main):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # ================================================================================
        # Preset
        # Button's Icon
        # self.setWindowIcon(QIcon('resource/img/ADU.svg'))
        self.setWindowTitle("출결확인서 자동 작성 도구")

        self.fileButton.setIcon(QIcon('resource/img/document-open-symbolic.svg'))
        self.chromeButton.setIcon(QIcon('resource/img/google-chrome.svg'))
        self.msedgeButton.setIcon(QIcon('resource/img/microsoft-edge.svg'))

        self.fileButton.clicked.connect(self.table_file_load)
        self.runButton.clicked.connect(self.run_making_pdf)

    def table_file_load(self):
        # return data : ('[file path]','file type')
        ori_file_path, _ = QFileDialog.getOpenFileName(
            self,
            "출결 파일 선택",
            "./",
            "Excel(*.xlsx *xls)"
        )

        if ori_file_path != "":
            global path_csv
            global path_directory
            ori_file_name = ori_file_path.split('/')[-1]

            # path/date_FileName
            path_directory = "%s/%s_%s" % (
                ori_file_path[0:ori_file_path.rfind('/')],
                time.strftime("%Y%m%d_%H%M", time.localtime(time.time())),
                ori_file_name.split('.')[0]
            )

            if os.path.isdir(path_directory) == 1:
                print("error")
            else:
                os.mkdir(path_directory)
                os.mkdir("%s/%s" % (path_directory, "raw"))
                os.mkdir("%s/%s" % (path_directory, "pdf"))

                shutil.copy("./resource/form_style.css", "%s/%s/form_style.css" % (path_directory, "raw"))

                # xlsx to csv
                xlsx = pandas.read_excel(ori_file_path)
                path_csv = "%s/%s.csv" % (path_directory, ori_file_name.split(".")[0])
                xlsx.to_csv(path_csv)

    def run_making_pdf(self):
        global path_csv
        global path_directory

        teacher = self.lineEdit.text()
        tmp_grade = self.spinBox.value()
        tmp_class = self.spinBox_2.value()

        self.statusBar.showMessage("doing")

        # 학생 출결 정보 파일을 불러옴
        with open(path_csv, 'r', encoding='UTF-8') as table:

            # 한 줄씩 처리
            for n in table.readlines()[1:]:
                # variable
                tmp_type = n.split(",")[4]

                if tmp_type[-2:] == "결석" or tmp_type == "출석인정조퇴":
                    if tmp_type != "미인정결석":
                        print(n, end="")

                        _, date, number, name, _, reason = n.split(",")
                        year, month, day = date.split(".")

                        path_raw_file = "%s/raw/%s%s%s_%s.html" % (path_directory, year, month, day, name)
                        with open(path_raw_file, 'w', encoding='UTF-8') as output:

                            # load resource file
                            with open("./resource/form.html", 'r', encoding='UTF-8') as form:
                                tmp_data = form.read()

                            # check absent type
                            if tmp_type == "출석인정결석":
                                codeA = "Ｏ"; codeB = ""; codeC = ""
                            elif tmp_type == "질병결석":
                                codeA = ""; codeB = "Ｏ"; codeC = ""
                            elif tmp_type == "출석인정조퇴":
                                codeA = ""; codeB = ""; codeC = "Ｏ"

                            # input data
                            old_data = ["$yr", "$mo", "$dy",
                                        "$grade", "$class", "$number", "$name",
                                        "$reason", "$state", "$teacher",
                                        "$a", "$b", "$c"
                                        ]

                            new_data = [year, month, day,
                                        tmp_grade, tmp_class, number, name,
                                        reason, tmp_type[-2:], teacher,
                                        codeA, codeB, codeC
                                        ]

                            for i in range(0,len(old_data)):
                                tmp_data = tmp_data.replace(str(old_data[i]), str(new_data[i]))
                            output.write(tmp_data)

                        if self.chromeButton.isChecked:
                            program = path_chrome
                        elif self.msedgeButton.isChecked:
                            program = path_msedge

                        os.system("%s --headless --print-to-pdf-no-header --print-to-pdf=%s --no-margins %s" % (
                            program, "%s/pdf/%s%s%s_%s.pdf" % (path_directory, year, month, day, name), path_raw_file
                        ))

        self.statusBar.showMessage("done")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = WindowMainClass()
    mainWindow.show()
    app.exec_()
