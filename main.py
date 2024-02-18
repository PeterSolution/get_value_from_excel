import openpyxl
from tkinter import filedialog
from tkinter import Tk

root = Tk()
root.withdraw()  # Ukryj główne okno

file_path = filedialog.askopenfilename(
    title="select file",
    filetypes=[("Excel", "*.xlsx"), ("Word", "*.docx"), ("all files", "*.*")]
)
excelheight=0
excelwidth=0
truedatah=0
truedatahw=0

checkifemplty=0
checkifempltywidth=0
checkwidth=0
checkheight=0
forwidth=10000
forheight=10000
if file_path:
    print("file:", file_path)
    sstr = file_path.rfind('/')
    sub = file_path[sstr + 1:-1]

    excel = openpyxl.load_workbook(file_path)

    print("what we have:")
    for sheet_name in excel.sheetnames:
        print(sheet_name)

    try:
        name=excel.sheetnames
        for data in name:
            data = excel[data]
            for i in range(forwidth):

                for j in range(forheight):
                    value = data.cell(row=i + 1, column=j + 1).value
                    if value is not None:
                        print(f'value from cell ({i + 1},{j + 1}): {value}')
                        if excelwidth < j:
                            excelwidth = j+1
                        if excelheight < i:
                            excelheight = i+1
                        checkifemplty = 0
                        if (forheight - j) < 10:
                            forheight = forheight + 1
                    if value is None:
                        checkifemplty = checkifemplty + 1
                        if j < 11 & checkifemplty == 10:
                            checkifempltywidth = checkifempltywidth + 1
                        else:
                            if j > 11 & checkifemplty == 10:
                                break
                    #     if j > excelwidth:
                    #         excelwidth = j
                    # if i > excelheight:
                    #     excelheight = i
            #         print(value)
            #
            # print(sub)
            # print(type(sub))

    finally:
        excel.close()

# checkifemplty=0
# checkifempltywidth=0
# checkwidth=0
# checkheight=0
# forwidth=100
# forheight=100
# while checkifempltywidth==10:
#     excel = openpyxl.load_workbook(file_path)
#     name = excel.sheetnames
#     for data in name:
#         data = excel[data]
#         for i in range(forwidth):
#
#             for j in range(forheight):
#                 value = data.cell(row=i + 1, column=j + 1).value
#                 if value is not None:
#                     print(f'value from cell ({i + 1},{j + 1}): {value}')
#                     if excelwidth<j:
#                         excelwidth=j
#                     if excelheight<i:
#                         excelheight=i
#                     checkifemplty=0
#                     if (forheight-j)<10:
#                         forheight=forheight+1
#                 if value is None:
#                     checkifemplty=checkifemplty+1
#                     if j <10&checkifemplty==10:
#                         checkifempltywidth=checkifempltywidth+1
#                     else:
#                         if j >10&checkifemplty==10:
#                             break

print("height: " + str(excelheight))
print("width: " + str(excelwidth))
