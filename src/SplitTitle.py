import collections
from WriteExcel import WriteExcel as exl_writer
import openpyxl
from janome.tokenizer import Tokenizer


class SplitTitle:
    @staticmethod
    def main():
        # 読み込みエクセル
        book = openpyxl.load_workbook("タイトル一覧.xlsx")

        t = Tokenizer()
        # 空の辞書
        d1 = collections.Counter({})

        # 見出し用のヘッダー
        header = ["単語", "出現頻度"]

        active_sheet = book.active
        # 2000行読み込み
        for i in range(2, 2002):
            s = active_sheet.cell(column=2, row=i).value

            d2 = collections.Counter(token.base_form for token in t.tokenize(s)
                                     if token.part_of_speech.startswith("名詞"))
            # 同一keyでvalueを加算する
            d1 += d2

        # 書き込み用配列作成
        write_list = []
        for key, val in sorted(d1.items(), key=lambda x: -x[1]):
            row = [key, val]
            write_list.append(row)

        # エクセルに書き込み
        exl_writer.write_to_excel(value=write_list, filename="単語の出現頻度", header=header)


if __name__ == "__main__":
    SplitTitle.main()
