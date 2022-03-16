import os
import openpyxl


class WriteExcel:
    @staticmethod
    def write_to_excel(value, folder_path=os.getcwd(), filename="タイトル一覧"):
        wb = openpyxl.Workbook()
        ws = wb.active

        # 表の見出し項目
        ws.append(["順位", "タイトル", "リンク先", "文字数"])

        # リストをセルに書き込み

        for row in value:
            ws.append(row)

        wb.save(os.path.join(folder_path, filename + ".xlsx"))
