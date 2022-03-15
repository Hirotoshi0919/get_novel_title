import time
from selenium import webdriver
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome import service as fs


class GerNovelTitle:
    @staticmethod
    def main():
        # ドライバー指定でChromeブラウザを開く
        chrome_service = fs.Service(executable_path="../driver/chromedriver.exe")
        driver = webdriver.Chrome(service=chrome_service)
        driver.implicitly_wait(10)

        # Googleアクセス
        driver.get("https://yomou.syosetu.com/search.php?search_type=novel&word=&button=")

        time.sleep(1)

        # プルダウンで並び替え
        pull_down = driver.find_element(By.NAME, "order")
        select = Select(pull_down)
        select.select_by_value("yearlypoint")

        time.sleep(2)

        write_list = []
        # NEXTボタンを押しつつ、タイトル取得
        while len(driver.find_elements(By.CLASS_NAME, "nextlink")) > 0:
            next_button = driver.find_elements(By.CLASS_NAME, "nextlink")
            next_button[0].click()

            # タイトル取得
            title_list = driver.find_elements(By.CLASS_NAME, "tl")

            # タイトルと文字数、リンク先の多次元配列を作成
            for title in title_list:
                title.text
                url = title.get_attribute("href")
                row = [title.text, url, len(title.text)]
            write_list.append(row)

        # ブラウザを閉じる
        driver.quit()

        print(write_list)

if __name__ == "__main__":
    get_title = GerNovelTitle

    get_title.main()
