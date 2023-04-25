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

        self.Button_file.clicked.connect(self.load_table_file)
        self.Button_run.clicked.connect(self.make_html)

    def load_table_file(self):
        # return data : ('[file path]','file type')
        ori_file_path, _ = QFileDialog.getOpenFileName(
            self,
            "출결 파일 선택",
            "./",
            "Excel(*.xlsx *xls)"
        )
        print(ori_file_path)

        # make directory name
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

            # check if same directory is already existing
            if os.path.isdir(path_directory) == 1:
                i = 1

                while os.path.isdir("%s(%d)" % (path_directory, i)) == 1:
                    i += 1

                path_directory = "%s(%d)" % (path_directory, i)

            if (0 == make_result_directory(ori_file_path, ori_file_name)):
                self.statusBar.showMessage("%s is ready." % ori_file_name)

    def make_html(self):
        global path_csv
        global path_directory

        # check if csv file is existing
        if path_csv == "":
            self.statusBar.showMessage("Load file first")

        else:
            file_data = ""

            teacher = self.lineEdit.text()
            num_grade = self.SBox_grade.value()
            num_class = self.SBox_class.value()

            self.statusBar.showMessage("start")

            with open(path_csv, 'r', encoding='UTF-8') as table:  # 학생 출결 정보 파일을 불러옴

                for line in table.readlines()[1:]:  # Process one line by one line

                    date, number, name, type, reason = split_line_data(line)

                    if (1 == check_absent_type(type, reason)):
                        print(line, end="")

                        year, month, day = date.split(".")  # split date

                        absent_marking = ["　", "　", "　"]  # check absent type
                        index_of_absent = ["출석인정결석", "질병결석", "출석인정조퇴"].index(type)
                        absent_marking[index_of_absent] = "Ｏ"

                        name = (len(name) < 6) if ("　" * (5 - len(name)) + name) else name

                        if len(teacher) < 6:
                            teacher = "　" * (5 - len(teacher)) + teacher

                        # input data
                        old_data = ["$yr", "$mo", "$dy",
                                    "$grade", "$class", "$number", "$name",
                                    "$reason", "$state", "$teacher",
                                    "$a", "$b", "$c"
                                    ]

                        new_data = [year, month, day,
                                    num_grade, num_class, number, name,
                                    reason, type[-2:], teacher,
                                    absent_marking[0], absent_marking[1], absent_marking[2]
                                    ]

                        # load resource file
                        with open("./resource/form_mid.html", 'r', encoding='UTF-8') as form:
                            tmp_data = form.read()

                            for i in range(0, len(old_data)):
                                tmp_data = tmp_data.replace(str(old_data[i]), str(new_data[i]))

                            file_data += tmp_data
                            file_data += "<div style='page-break-before:always'></div>\n"

            # last page fix
            file_data = file_data[:-len("<div style='page-break-before:always'></div>\n")]

            # copy html <head> data to all.html from resource
            with open("%s/all.html" % path_directory, 'w', encoding='UTF-8') as file:
                with open("./resource/form.html", 'r', encoding='UTF-8') as html:
                    tmp_file_data = html.read() % file_data
                    file.write(tmp_file_data)

            self.statusBar.showMessage("done")


def make_result_directory(file_path, file_name):
    global path_csv
    global path_directory

    os.mkdir(path_directory)

    shutil.copy("./resource/form_style.css", "%s/form_style.css" % path_directory)

    # xlsx to csv
    xlsx = pandas.read_excel(file_path)
    path_csv = "%s/%s.csv" % (path_directory, file_name.split(".")[0])
    xlsx.to_csv(path_csv)

    return 0


def check_absent_type(type, reason):
    if type[-2:] == "결석" or type == "출석인정조퇴":
        if type != "미인정결석":
            if reason != "가정학습" and reason != "가정 학습":
                return 1

    return 0


def split_line_data(line):
    split_data = line[:-1].split(",")           # line[:-1] : delete "\n"

    if len(split_data) > 6:                     # check if reason included ","
        _, date, number, name, type = split_data[:5]
        reason = ', '.join(str(s) for s in split_data[5:])

        if reason[-1:] == '"':
            reason = reason[1:-1]               # delete "" that made when xlsx to csv
        else:                                   # when reason was had "\n" before convert csv
            reason = reason[1:]

        return date, number, name, type, reason

    elif len(split_data) == 6:
        _, date, number, name, type, reason = split_data
        return date, number, name, type, reason

    else:                                       # when previous reason had \n
        return 0, 0, 0, "0", "0"


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = WindowMainClass()
    mainWindow.show()
    app.exec_()
