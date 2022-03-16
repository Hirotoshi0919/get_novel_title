from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome import service as fs
from WriteExcel import WriteExcel as writer


class GerNovelTitle:

    @staticmethod
    def main():
        # ドライバー指定でChromeブラウザを開く
        chrome_service = fs.Service(executable_path="../driver/chromedriver.exe")
        driver = webdriver.Chrome(service=chrome_service)
        driver.implicitly_wait(10)

        try:

            # ホームページにGoogleでアクセス
            driver.get("https://yomou.syosetu.com/search.php?search_type=novel&word=&button=")

            # プルダウンで並び替え
            pull_down = driver.find_element(By.NAME, "order")
            select = Select(pull_down)
            select.select_by_value("yearlypoint")

            # エクセル書き込み用のリスト
            write_list = []
            rank = 1
            while True:
                # タイトル取得
                title_list = driver.find_elements(By.CLASS_NAME, "tl")

                # タイトルと文字数、リンク先の多次元配列を作成
                for title in title_list:
                    url = title.get_attribute("href")
                    row = [rank, title.text, url, len(title.text)]
                    write_list.append(row)
                    rank += 1

                if len(driver.find_elements(By.CLASS_NAME, "nextlink")) > 0:
                    next_button = driver.find_elements(By.CLASS_NAME, "nextlink")
                    next_button[0].click()
                else:
                    break

            writer.write_to_excel(write_list)

        except Exception as e:
            raise e

        finally:
            # ブラウザを閉じる
            driver.quit()


if __name__ == "__main__":
    get_title = GerNovelTitle
    get_title.main()
