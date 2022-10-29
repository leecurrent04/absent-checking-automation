import pandas as pd
import pdfkit

tmp_grade = 3
tmp_class = 3
# tmp_class = int(input("반 정보를 입력하세요: "))
teacher = "정성일"

# xlsx to csv
xlsx = pd.read_excel("./resource/table.xlsx")
xlsx.to_csv("./resource.csv")

# 학생 출결 정보 파일을 불러옴
with open("./resource.csv", 'r') as table:
    
    # 한 줄씩 처리
    for n in table.readlines()[1:]:
        
        # variable
        tmp_type = n.split(",")[4]

        if tmp_type[-2:] == "결석" or tmp_type=="출석인정조퇴":
            if tmp_type != "미인정결석":
                print(n,end="")

                _, date, number, name, _, reason = n.split(",")
                year, month, day = date.split(".")

                with open("./output/raw/%s%s%s_%s.html"%(year,month,day, name), 'w') as output: 
                    with open("./resource/form.html",'r') as form:
                        tmp_data = form.read()


                    if tmp_type == "출석인정결석":
                        codeA = "Ｏ"; codeB = ""; codeC = ""
                    elif tmp_type == "질병결석":
                        codeA = ""; codeB = "Ｏ"; codeC = ""
                    elif tmp_type == "출석인정조퇴":
                        codeA = ""; codeB = ""; codeC = "Ｏ"

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

                path_input = "./output/raw/%s%s%s_%s.html"%(year,month,day, name)
                path_output = "./output/%s%s%s_%s.pdf"%(year,month,day, name)
                options= {
                        'page-size': 'A4',
                        'margin-top': '0',
                        'margin-right': '0',
                        'margin-bottom': '0',
                        'margin-left': '0',
                        'disable-smart-shrinking': "",
                        "enable-local-file-access": ""
                }
                path_css = "./output/raw/form_style.css"
                pdfkit.from_file(path_input, path_output, options=options, css=path_css)
