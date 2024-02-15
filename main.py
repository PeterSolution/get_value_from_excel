import openpyxl
from tkinter import filedialog
from tkinter import Tk

root = Tk()
root.withdraw()  # Ukryj główne okno

file_path = filedialog.askopenfilename(
    title="select file",
    filetypes=[("Excel", "*.xlsx"), ("Word", "*.docx"), ("all files", "*.*")]
)

if file_path:
    print("file:", file_path)
    # Tutaj możesz dodać kod obsługujący wybrany plik
    str = file_path.rfind('/')
    sub = file_path[str + 1:-1]

    excel = openpyxl.load_workbook(file_path)

    # Wydrukuj dostępne arkusze w pliku
    print("what we have:")
    for sheet_name in excel.sheetnames:
        print(sheet_name)

    try:
        # Spróbuj uzyskać dostęp do arkusza 'GetData'
        name=excel.sheetnames
        for data in name:
            data = excel[data]

            # Dalsza część kodu
            for i in range(10):
                for j in range(10):
                    value = data.cell(row=i + 1, column=j + 1).value
                    if value is not None:
                        print(f'value from cell ({i + 1},{j + 1}): {value}')

            #         print(value)
            #
            # print(sub)
            # print(type(sub))

    finally:
        # Zamyka plik Excel
        excel.close()
