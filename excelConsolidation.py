from pathlib import Path
import time
import xlwings as xw
import os


def consolidate():
    SOURCE_DIR = os.getcwd()
    print(SOURCE_DIR)
    print(list(Path(SOURCE_DIR).glob("Comparison*.xlsx")))

    excel_files = list(Path(SOURCE_DIR).glob("Comparison*.xlsx"))
    combined_wb = xw.Book()
    t = time.localtime()
    timestamp = time.strftime("%Y-%m-%d_%H%M", t)

    for excel_file in excel_files:
        wb = xw.Book(excel_file)
        for sheet in wb.sheets:
            sheet.api.Copy(After=combined_wb.sheets[0].api)
        wb.close()

    try:
        combined_wb.sheets[0].delete()
    except:
        pass
    combined_wb.save(f'MasterFile_{timestamp}.xlsx')
    if len(combined_wb.app.books) == 1:
        combined_wb.app.quit()
    else:
        combined_wb.close()

    for excel_file in excel_files:
        os.remove(excel_file)


if __name__ == "__main__":
    consolidate()