
# import IRpwBB4        (キャッシュ処理インポート版)
#
# irBB = IRpwBB2.IRBankBrowser()
# irBB.start()      #初期化 (playwriteブラウザ起動)
# irBB.cleanup()    #終了処理
# irBB.stock_code   #銘柄コードの設定 ex. irBB.stock_code = "2708"
#株式情報ページ (例: https://irbank.net/1452 )から 内部コードを取得できないものは処理対象外

# irBB.stock_code                       # 銘柄コードの設定先
# irBB.cache_cleared
# irBB.remove_ads()                     # 広告削除（銘柄ごとに独立）
# irBB.load_pl_page()                   # 会社業績ページ(/pl)読込
# irBB.copy_sales_graph()               #   売上高グラフ（SVG #0）
# irBB.copy_net_income_graph()          #   当期純利益グラフ（SVG #2）
# irBB.load_per_page()                  # PER推移ページ(/per)読込
# irBB.copy_per_graph()                 #   PERグラフ
# dfV  = irBB.read_value_table()        # 価値算定テーブル(/value)読込
# 一覧表列名: '年度', '財産価値', '株主価値', '1株価値', '株価', '時価総額'
# dfCF = irBB.read_cf_table()           # CF推移テーブル(/cf)を取得
# 一覧表列名: "年度", "単位", "四半期", "営業CF", "投資CF", "財務CF", "フリーCF", "設備投資", "現金等"
# (単位) 1:百万円、100:億円

from IRCache import CsvManager, CacheManager
from IRCache import CSV_COLNAME, CSVCN1, CSVCN2, CSVCN3, CSVCN4, CSVCN5
# IRCODE.csv を事前に準備する。IRCache.py 冒頭の説明参照。

#from bs4 import BeautifulSoup

from playwright.sync_api import sync_playwright, TimeoutError # py -m pip install playwright
# pip installの後、Scriptsフォルダで playwright install 
# Scriptsフォルダは py /? してpython.exeの場所を確認。そのサブフォルダである。

from PIL import Image                   # py -m pip install pillow
from io import BytesIO
import time
import re
import pandas as pd                     # py -m pip install pandas
from datetime import datetime, date


#########################################################################
import win32clipboard                   # py -m pip install pywin32
import win32con

def copy_image_to_clipboard(image):         # imageをクリップボードにコピー
    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    win32clipboard.CloseClipboard()

    image.close()


#########################################################################
class IRBankBrowser:
    def __init__(self):
        # --- 新しいキャッシュクラス ---
        print("(IRBankBrowser:__init__) load cache ...", end="")
        self.csv = CsvManager("IRCODE.csv", CSV_COLNAME)# CSV読込
        self.cache = CacheManager("irbank_cache.json")  # キャッシュ読込
        print("ok")
        
        self.playwright = None
        self.browser = None
        self.page = {}          # self._urlに対応する各page
        self.svgs = None        # 処理中tagのlocatorの値

        self._stock_code = None         # 銘柄コード
        self._cache_cleared = False     # キャッシュがクリアされた
        self._stock_name = None         # 銘柄名

        self._url_base = 'https://irbank.net/'
        self._url = {}          # 読込対象 url
        self._ref_list = []     # 先行読み込みページリスト 例: ["value","cf","per"]


    # --- with 用 ---
    def __enter__(self):
        self.start()
        return self
    def __exit__(self, exc_type, exc, tb):
        self.cleanup()

    # --- with なし用 ---
    def start(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=True)
        print("(start)ブラウザ起動完了")
        
    def cleanup(self):
        self.csv.update_row()   # CSV 保存
        self.cache.save()       # json保存
        for key in self.page.keys():
            self.page[key].close()
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        print("(cleanup)ブラウザを閉じました")

    @property
    def stock_code(self):           # 銘柄コード読み出し
        return self._stock_code

    @stock_code.setter
    def stock_code(self, value):    # 銘柄コード設定

        if self._cache_cleared:     # 前回キャッシュクリアされた(Web読込した)
            self.csv.update_row()# CSV保存
            self.cache.save()    # json保存

        self._stock_code = value    # 銘柄コード
        
        self.cache.stock_code = value
        self.csv.stock_code = value
        
        if not self.csv.row:
            print(f"\n######## 無効な銘柄コード: {value} ########\n")
            self._stock_code = None
            # csv.stock_codeは csv側でNoneを設定、cache.stock_codeはそのまま
            return
        else:
            print(f"銘柄コード: {value}")

        for key, page in self.page.items(): #開いているページを閉じる
            if page:
                try:
                    page.close()
                except:
                    pass
        self.page = {}
        
        self._stock_name = self.csv.row[ CSV_COLNAME[ CSVCN2 ] ] # 銘柄名

        # URL 設定 #1
        # top:株式情報ページ (ir_code取得目的)
        # per:PER推移ページ
        self._url["top"] = f"{self._url_base}{self._stock_code}/"
        ir = self.ir_code()         # IRBANKコード
        if not ir:          #無効なIRBANKコード
            print(f"(stock_code)IRBANKコード取得不可: {value}")
            self._stock_code = None
            return
        self._url["per"] = f"{self._url['top']}per?mw=2"
        
        # URL 設定 #2
        # ir:会社情報ページ
        # pl:会社業績ページ  value:価値算定ページ  cf:CF推移ページ
        self._url["ir"]    = f"{self._url_base}{ir}/"
        self._url = self._url | {key: f"{self._url['ir']}{key}"
            for key in ("pl", "value", "cf")}

        self._cache_cleared = self.check_announcement()    #直近決算発表日が更新されていたらキャッシュクリア
        if self._cache_cleared:
            self.prefetch_pages()   # self._ref_list のページを並行取得

    @property
    def stock_name(self):
        return self._stock_name
    @property
    def cache_cleared(self):
        return self._cache_cleared

    @property
    def ref_list(self):
        return self._ref_list
    @ref_list.setter
    def ref_list(self, value):    # 先行読込ページリスト
        self._ref_list = value
    
    @property
    def url_cf(self):
        return self._url["cf"]
    @property
    def url_value(self):
        return self._url["value"]
    @property
    def url_per(self):
        return self._url["per"]
    @property
    def url_pl(self):
        return self._url["pl"]
    @property
    def next_announcement(self):  #次回決算発表日
        return self.csv.get(CSV_COLNAME[CSVCN5]) 
    @property
    def announcement(self):       #決算発表日
        return self.csv.get(CSV_COLNAME[CSVCN4]) 

    ##########################################################################################
    def get_ir_code(self) -> str:# 銘柄トップページからIRBANKコードを取得
                                            # ir_code_update()から呼び出す
        key = "top"
        self.open_page(key)

        # <ul class="nsq"> を全部取得        #画面左領域のリンク集
        uls = self.page[key].locator("ul.nsq")
        count = uls.count()

        if count == 0:  # URL誤りもここに来る
            print(f"(get_ir_code){self.stock_code}: ul.nsq が見つかりません")
            a_tag = self.page[key].locator('a[title="有価証券（EDINET）"]')    #1452には<ul class="nsq">がない
            if a_tag.count() != 0:
                print(f"(get_ir_code){self.stock_code}: 有価証券（EDINET）リンクを参照します")
                # count()==0のときは関数末尾まで処理なしで到達する

        else:   # 各 ul の中の <a title="有価証券報告書"> を探す
            for i in range(count):
                ul = uls.nth(i)
                #a_tag = ul.locator('a[title="有価証券報告書"]')  
                a_tag = ul.locator('a[title*="有価証券"]')

                if a_tag.count() != 0:
                    break

        if a_tag.count() != 0:
            href = a_tag.first.get_attribute("href")  # 例: "/E02938/edinet"
            if href:
                parts = href.split("/")
                # ["", "E02938", "edinet"]

                if len(parts) >= 2 and parts[1] and parts[1][0].isalpha():
                    internal_id = parts[1]
                    print(f"(get_ir_code){self.stock_code}: ir_code = {internal_id}")
                    return internal_id
        
        print(f"(get_ir_code){self.stock_code}: 内部IDを含むリンクが見つかりませんでした")
        return None

    def ir_code(self):                   # IRBANKの内部コードを取得して返す。
        ir = self.csv.get(CSV_COLNAME[CSVCN3])
        if not ir:      #IRコード　未取得
            ir = self.get_ir_code()       # IRBANKのWebページから読み取る
            if ir:      #IRコード取得 成功
                self.csv.set(CSV_COLNAME[CSVCN3], ir)

        print("(ir_code) ", self.stock_code, "-->", ir)
        return ir
            
    ##########################################################################################
    def prefetch_pages(self):       #self._ref_list 指定ページを並行取得
        if not hasattr(self, "_ref_list"):
            return
        for key in self._ref_list:
            self.open_page(key)
        print(f"(prefetch_pages) {self._ref_list} を並行読み込み開始しました")

    def open_page(self, key):       # key指定ページをWeb読込
        if self.page.get(key):  #すでに開いていたら、開いているものを使う
            return
        # ページを開く
        page = self.browser.new_page()
        response = page.goto(self._url[key]) #, wait_until="domcontentloaded")
    # response is None: DNSエラー、response.status >= 400: HTTPエラー
        self.page[key] = page
    #（Chromium 内部では非同期で読み込みが進み、locator処理時に要素が見つかるまで waitする)

    def find_optional(self, page, selector, timeout=3000):  #無い場合もある場合の locatorを返す
        locator = page.locator(selector)
        try:
            locator.first.wait_for(timeout=timeout)
            return locator
        except TimeoutError:
            return None

    ##########################################################################################
    def get_latest_announcement_datetime(self): # 直近決算発表日時を取得
        key = "pl"          # PLページ
        self.open_page(key)

        # <p class="pad"> の最初の <a> の text が決算発表日時
        locator = self.find_optional(self.page[key], "p.pad a")
        if locator:
            return locator.inner_text().strip() # 例: "2026年2月25日 15:40"
        else:
            return None

    def get_next_announcement_date(self):       # 次回決算発表日を取得
        key = "top"         #株式情報ページから「次の決算発表日」を取得
        self.open_page(key)

        # 「次の決算発表は◯月◯日を予定しています。」の <a> を取得
        locator = self.find_optional(self.page[key],
                            'div.message_n a[title*="決算スケジュール"]')
        if not locator:
            return None

        # text = locator.inner_text().strip()     # 例: '5月7日'
        # 年はページの URL から取得できないので、リンクの title から取る
        title = locator.get_attribute("title")
                                    # 例: '決算スケジュール - 2026年5月7日'

        # 年月日の抽出
        m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", title)
        if not m:
            return None
        y, mo, d = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"#例: '5月7日' → '2026-05-07'
    
    def should_skip_pl_fetch(self, next_dt):    #次回決算発表日が未来なら True
        next_date = datetime.strptime(next_dt, "%Y-%m-%d").date()
        return next_date > date.today()

    def is_next_announce_past_or_same(self, cached_next, latest):
        # cached_next が latest の日付と同じか過去なら True
        # cached_next（次回決算発表日: 'YYYY-MM-DD'）
        # latest（直近決算発表日時: '2026年2月25日 15:40'）

        # latest を datetime に変換
        latest_dt = datetime.strptime(latest, "%Y年%m月%d日 %H:%M")

        # cached_next を datetime に変換（時刻は 00:00 とみなす）
        next_dt = datetime.strptime(cached_next, "%Y-%m-%d")

        # 次回決算発表日が latest と同じ or 過去なら True
        return next_dt <= latest_dt
    
    def check_announcement(self):
        # PLページの直近決算発表日が更新されていたらキャッシュをクリアする
        # 但し、次回決算発表日が未来なら処理しない
        # (戻り値) True:キャッシュをクリアした False:してない

        cached_next = self.csv.get(CSV_COLNAME[CSVCN5]) #キャッシュ読込 (次回決算発表日)
        if not cached_next:                         #キャッシュにない
            next_dt = self.get_next_announcement_date() #Webから取得 (次回決算発表日)
            if not next_dt:                         #取得できない
                print(f"(check_announcement) 次回決算発表日時　取得失敗(1)")
            else:
                self.csv.set(CSV_COLNAME[CSVCN5], next_dt)  #キャッシュ保存 (次回決算発表日)
        else:
            next_dt = None  #Webから未取得

        cached = self.csv.get(CSV_COLNAME[CSVCN4])      #キャッシュ読込 (決算発表日)
        if cached and (cached_next or next_dt) and self.should_skip_pl_fetch( cached_next or next_dt ):
            # 決算発表日がキャッシュにあり次回決算発表日が未来なら、PL を読みに行かない
            print(f"(check_announcement) 次回決算発表日 {cached_next} が未来のため、PLページの直近決算発表日を読みに行きません")
            return False

        latest = self.get_latest_announcement_datetime()    #Webから取得 (決算発表日)
        if not latest:
            print(f"(check_announcement) 直近決算発表日 取得失敗")

        if cached != latest:                                # 決算発表日 更新あり
            print(f"(check_announcement) 決算発表日時が更新されました: {cached} → {latest}")

        if cached != latest or not latest:  #直近決算発表日が更新されてる or 取得できない
            self.cache.clear()                          #キャッシュクリア
            self.csv.set(CSV_COLNAME[CSVCN4], latest)   #キャッシュ保存 (決算発表日)
            self.csv.set(CSV_COLNAME[CSVCN5], next_dt)  #キャッシュ保存 (次回決算発表日)
            return True
        else:                               #直近決算発表日が更新されていない
            print(f"(check_announcement) 決算発表日時は変更なし: {latest}")
            
            if cached_next and self.is_next_announce_past_or_same(cached_next, latest):
                print(f"(check_announcement) 次回決算発表日 {cached_next} は latest {latest} と同じか過去のため、無効化します")
                self.csv.set(CSV_COLNAME[CSVCN5], None)
                
            return False


    ##########################################################################################
    def remove_ads(self, key):                  # ページの広告を削除
        page = self.page[key]
        page.evaluate("""
            document.querySelectorAll('*').forEach(el => {
                if (getComputedStyle(el).position === 'fixed') el.remove();
            });
        """)
        print(f"(remove_ads)広告を削除しました({key})")

    def set_page(self, html, tag):              #HTMLをplaywriteブラウザに読み込む
        key = "html"
        if self.page.get(key):
            self.page[key].close()        
        self.page[key] = self.browser.new_page()

        html = f"<html><body>{html}</body></html>"""
        print("(set_page)Cacheページ読込")
        self.page[key].set_content(html)

        self.svgs = self.page[key].locator(tag)
        self.svgs.first.wait_for()

    def cache_set_svgs(self, key, tag=None):    # self.svgsをキャッシュに保存
        if tag:
            key += "-" + tag
        if not self.svgs:
            self.cache.set(key, None)
            return
        html = ""
        for i in range(self.svgs.count()):
            #el = self.svgs.nth(i).element_handle()
            #html = "{}<{}>{}</{}>".format(html, tag, el.inner_html(), tag)
            #⇒ inner_html()だと <svg のサイズが保存できない
            #html += el.evaluate("node => node.outerHTML")
            html += self.svgs.nth(i).evaluate("node => node.outerHTML")
        self.cache.set(key, html)

    def load_svgs(self, key, tag, selopt=""):       # page[key]の tagを self.svgsに設定する
        key2 = key + "-" + tag      # 新しいキャッシュキー    
        cached = self.cache.get( key2 )
        if not cached and (cached := self.cache.get( key )):    #key2は無いが keyはある
            self.cache.rename_key( key, key2 )  #keyをkey2に変更しておく
        # cached が "", "0", " ", "{}" 等の場合も not cachedは Trueとなる。(無問題)
            
        if cached:                                              #キャッシュデータ有り
            self.set_page( cached, tag )    # cacedを svgsに設定
            
            if selopt:
                selector = tag + selopt
                self.svgs = self.find_optional( self.page["html"], selector )
                if not self.svgs:
                    cached = None
            
        if not cached:                                          #キャッシュデータ無し
            print(f"(load_svgs) {self._url[key]}")
            self.open_page(key)
            self.svgs = self.find_optional( self.page[key], tag )

            if tag in ("svg"):
                self.copy_graph_first("広告付き")
                self.remove_ads(key)           #広告削除

            self.cache_set_svgs(key2)       # svgsを Cache保存

            if selopt:
                selector = tag + selopt
                self.svgs = self.find_optional( self.page[key], selector )

    def copy_graph_specified(self, name, cond):    # condのうち１つを含む<svg>をコピーする
        if not self.svgs:
            return False
        count = self.svgs.count()
        for i in range(count):
            svg = self.svgs.nth(i)
            html = svg.inner_html()
            if any( c in html for c in cond):
                png = svg.screenshot()
                img = Image.open(BytesIO(png))
                copy_image_to_clipboard(img)
                print(f"{name}グラフをコピーしました")
                time.sleep(0.4)
                return True
        return False

    def copy_graph_first(self, name):       # svgsのうち一つ目の<svg>をコピーする        
        if not self.svgs:
            return False
        png_bytes = self.svgs.first.screenshot()
        img = Image.open(BytesIO(png_bytes))
        copy_image_to_clipboard(img)
        print(f"{name}グラフをコピーしました")
        time.sleep(0.4)
        return True
    
    ##########################################################################################
    def load_pl_page(self):                 # 会社業績ページ(/pl)読込
        self.load_svgs( "pl", "svg", '[aria-label="グラフ。"]' )

    def copy_sales_graph(self, n="売上高"):    # 売上高グラフ（SVG #0）
        return self.copy_graph_specified("売上高", ["売上高", "包括利益", "営業収益"] )
        # "売上高"がないものもある ex) 4597, 8393

    def copy_net_income_graph(self):        # 利益率グラフ（SVG #1 or #2）
        return self.copy_graph_specified("利益率", ["税引前利益"] )

    ##########################################################################################
    def load_per_page(self):    # PER推移ページ(/per)読込
        self.load_svgs( "per", "svg", '[aria-label="グラフ。"]' )
        #2026/4/24 リファインバースグループ（7375） PERグラフ壊れてた Data column(s) for axis #0 cannot be of type string

    def copy_per_graph(self):               # PERグラフ
        return self.copy_graph_first("PER")

    ##########################################################################################
    # 価値算定ページを読み込み、テーブルを DataFrame で返す
    def read_value_table(self) -> pd.DataFrame:
        self.load_svgs( "value", "table", '.bar:has(caption:text("価値算定"))' )
        if not self.svgs:
            return pd.DataFrame()

        table = self.svgs

        ths = table.locator("thead tr th")      # --- 列名（thead） ---
        colnames = [ths.nth(i).inner_text().strip() for i in range(ths.count())]

        trs = table.locator("tbody tr")         # --- データ行（tbody） ---
        data = []

        for i in range(trs.count()):
            tr = trs.nth(i)
            tds = tr.locator("td")

            row = []
            for j in range(tds.count()):
                # 1行目だけ取得（IRBANK は <span> が複数ある）
                raw = tds.nth(j).inner_text().split("\n")[0].strip()
                row.append(raw)

            # 年度 YYYY/MM → YYYYMM
            row[0] = re.sub(r"/", "", row[0])

            data.append(row)

        df = pd.DataFrame(data, columns=colnames)
        df.set_index(colnames[0], inplace=True)

        return df

    ##########################################################################################
    # CF推移テーブルを取得
    def read_cf_table(self) -> pd.DataFrame:
        self.load_svgs( "cf", "table", '.cs' )
        if not self.svgs:
            return pd.DataFrame()
            # CF テーブル   #2026/4/25 「横浜ライト工業のキャッシュ・フローの推移に関する情報は見つかりませんでした。」

        table = self.svgs

        caption_text = table.locator("caption").inner_text()    # 単位判定（caption に "億円" が含まれるか）
        unit = 100 if "億円" in caption_text else 1

        ths = table.locator("thead tr th")                      # --- 列名（thead） ---
        colnames = []
        for i in range(ths.count()):
            name = ths.nth(i).inner_text().strip()
            name = re.sub(r"#\d+$", "", name)   # 営業CF#1 → 営業CF
            colnames.append(name)

        colnames.insert(1, "単位")            # 列名に「単位」を追加

        trs = table.locator("tbody tr")                         # --- データ行（tbody） ---
        n = trs.count()

        data = []
        i = 0

        while i < n:
            tr = trs.nth(i)
            tds = tr.locator("td")
            td_count = tds.count()

            # 年度セル（rowspan がある行） (1行しかなくてもrowspan=1はある)
            rowspan = tds.nth(0).get_attribute("rowspan")
            if rowspan and rowspan.isdigit():
                rowspan = int(rowspan)
            else:
                i += 1
                continue

            # 年度 YYYY年MM月期 → YYYYMM
            raw_year = tds.nth(0).inner_text().replace("\n", "")
            m = re.match(r"(\d+)年(\d+)月", raw_year)
            if m:
                year = f"{m.group(1)}{int(m.group(2)):02d}"
            else:
                year = raw_year.strip()

            # rowspan 行のうち最後の行（四半期/通期）を読む
            full_row_index = i + (rowspan - 1)
            if full_row_index >= n:
                break

            full_tr = trs.nth(full_row_index)
            full_tds = full_tr.locator("td")

            # 最終四半期のデータ（四半期列を含む）
            row = []
            for j in range(full_tds.count()):
                val = full_tds.nth(j).inner_text().split("\n")[0].strip()
                row.append(val)
                
            # 1行しかない時は row[0]に年度列データがあるので削除
            if rowspan == 1:
                row.pop(0)
                
            # 年度 + 単位 + 通期データ
            data.append([year, unit] + row)

            # 次の年度へ
            i += rowspan

        # --- DataFrame 化 ---
        df = pd.DataFrame(data, columns=colnames)
        df.set_index(colnames[0], inplace=True)

        return df


#############################################################################################
# 実行 (テストドライバ)
if __name__ == "__main__":
    if 0 :                          # IRBANK内部コード取得
        with IRBankBrowser() as irBB:
            irBB.stock_code = "1301"
            irBB.stock_code = "2708"

    elif 1 :                          # IRBANK内部コード取得
        with IRBankBrowser() as irBB:
            irBB.stock_code = "6232"
            irBB.cache.clear()
            
    elif 0 :
        with IRBankBrowser() as irBB:
            irBB.stock_code = "2708"
            irBB.cache.clear()
            
    elif 0 :                        # 売上高グラフ・利益率グラフのコピー
        with IRBankBrowser() as irBB:
            irBB.stock_code = "2708"
            irBB.load_pl_page()         # 会社業績ページ(/pl)読込
            irBB.copy_sales_graph()     # 売上高グラフ（SVG #0）
            irBB.copy_net_income_graph()# 当期純利益グラフ（SVG #2）
            
    elif 0 :                        # PERグラフのコピー
        with IRBankBrowser() as irBB:
            irBB.stock_code = "2708"
            irBB.load_per_page()        # PER推移ページ(/per)読込
            irBB.copy_per_graph()       # PERグラフ
        
    elif 0 :                        # 価値算定テーブル
        import pyperclip                #pip install pyperclip
        with IRBankBrowser() as irBB:
            irBB.stock_code = "2708"
            dfV = irBB.read_value_table()
        tsv = dfV.to_csv(sep="\t") #, index=False)
        # 改行コードを強制的に \n に統一
        tsv = tsv.replace("\r\n", "\n").replace("\r", "\n")
        print(tsv, end="")
        pyperclip.copy(tsv) #クリップボードにコピー

    elif 0 :                        # CF推移テーブル
        import pyperclip                #pip install pyperclip
        with IRBankBrowser() as irBB:
            #irBB.stock_code = "2708"
            irBB.stock_code = "3659"
            dfCF = irBB.read_cf_table()
        tsv = dfCF.to_csv(sep="\t") #, index=False)
        # 改行コードを強制的に \n に統一
        tsv = tsv.replace("\r\n", "\n").replace("\r", "\n")
        print(tsv, end="")
        pyperclip.copy(tsv) #クリップボードにコピー


    else:
        irBB = IRBankBrowser()
        irBB.start()

        irBB.stock_code = "1301"
        irBB.stock_code = "2708"

        irBB.cleanup()


