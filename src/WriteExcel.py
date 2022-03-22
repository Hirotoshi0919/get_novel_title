import os
import openpyxl


class WriteExcel:
    @staticmethod
    def write_to_excel(value, header, folder_path=os.getcwd(), filename="タイトル一覧"):
        wb = openpyxl.Workbook()
        ws = wb.active

        # 表の見出し項目
        ws.append(header)

        # リストをセルに書き込み

        for row in value:
            ws.append(row)

        wb.save(os.path.join(folder_path, filename + ".xlsx"))
