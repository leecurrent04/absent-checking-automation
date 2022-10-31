import pandas
import time
import sys
import os
import os.path
import shutil
import platform

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


class WindowMainClass(QMainWindow, form_main):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # ================================================================================
        # Preset
        # Button's Icon
        # self.setWindowIcon(QIcon('resource/img/ADU.svg'))
        self.setWindowTitle("출결확인서 자동 작성 도구")

        self.Button_file.setIcon(QIcon('resource/img/document-open-symbolic.svg'))

        self.Button_file.clicked.connect(self.table_file_load)
        self.Button_run.clicked.connect(self.run_making_pdf)

    def table_file_load(self):
        # return data : ('[file path]','file type')
        ori_file_path, _ = QFileDialog.getOpenFileName(
            self,
            "출결 파일 선택",
            "./",
            "Excel(*.xlsx *xls)"
        )
        print(ori_file_path)

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

                shutil.copy("./resource/form_style.css", "%s/form_style.css" % path_directory)

                # xlsx to csv
                xlsx = pandas.read_excel(ori_file_path)
                path_csv = "%s/%s.csv" % (path_directory, ori_file_name.split(".")[0])
                xlsx.to_csv(path_csv)

            self.statusBar.showMessage("%s is ready." % ori_file_name)

    def run_making_pdf(self):
        global path_csv
        global path_directory
        file_data = ""

        teacher = self.lineEdit.text()
        tmp_grade = self.SBox_grade.value()
        tmp_class = self.SBox_class.value()

        self.statusBar.showMessage("start")

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

                        # load resource file
                        with open("./resource/form_mid.html", 'r', encoding='UTF-8') as form:
                            tmp_data = form.read()

                        # check absent type
                        if tmp_type == "출석인정결석":
                            codeA = "Ｏ"; codeB = ""; codeC = ""
                        elif tmp_type == "질병결석":
                            codeA = ""; codeB = "Ｏ"; codeC = ""
                        elif tmp_type == "출석인정조퇴":
                            codeA = ""; codeB = ""; codeC = "Ｏ"

                        if len(name) < 6 :
                            tmp_name = "　"*(5-len(name)) + name
                        else:
                            tmp_name = name

                        if len(teacher) < 6 :
                            teacher = "　"*(5-len(teacher)) + teacher

                        # input data
                        old_data = ["$yr", "$mo", "$dy",
                                    "$grade", "$class", "$number", "$name",
                                    "$reason", "$state", "$teacher",
                                    "$a", "$b", "$c"
                                    ]

                        new_data = [year, month, day,
                                    tmp_grade, tmp_class, number, tmp_name,
                                    reason, tmp_type[-2:], teacher,
                                    codeA, codeB, codeC
                                    ]

                        for i in range(0,len(old_data)):
                            tmp_data = tmp_data.replace(str(old_data[i]), str(new_data[i]))

                        file_data += tmp_data
                        file_data += "<div style='page-break-before:always'></div>\n"

        self.statusBar.showMessage("done")

        with open("%s/all.html" % path_directory, 'w', encoding='UTF-8') as file:
            with open("./resource/form.html", 'r', encoding='UTF-8') as html:
                tmp_file_data = html.read() % file_data
                file.write(tmp_file_data)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = WindowMainClass()
    mainWindow.show()
    app.exec_()
