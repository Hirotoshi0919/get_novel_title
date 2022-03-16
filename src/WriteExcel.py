import os
import openpyxl


class WriteExcel:
    @staticmethod
    def write_to_excel(val, folder_path=os.getcwd(), filename="タイトル一覧"):
        """
        エクセル書き込み用メソッド
        :param val: 書き込む値（2次元配列）
        :param folder_path: 出力先のパス（省略可）
        :param filename: 出力するエクセルのファイル名
        """
        wb = openpyxl.Workbook()
        ws = wb.active

        # 表の見出し項目
        ws.append(["順位", "タイトル", "リンク先", "文字数"])

        # リストをセルに書き込み

        for row in val:
            ws.append(row)

        wb.save(os.path.join(folder_path, filename + ".xlsx"))
