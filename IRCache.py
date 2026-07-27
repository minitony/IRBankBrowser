
# import IRCache
#
# IRcsv     = IRCache.CsvManager("IRCODE.csv", CSV_COLNAME) #
# IRcsv.stock_code      # 銘柄コードの指定
# IRcsv.update_row()    # 属性情報をCSVファイルに保存
# IRcsv.row[ CSV_COLNAME[CSVCN2] ]  #属性情報の取得元・設定先
# IRcsv.marge_stock_code() #銘柄コードの追加
# IRcsv.copy_past_next_announcement_codes() #次回決算発表日が過去のもの抽出

CSV_COLNAME = ["コード", "銘柄名", "IRBANKコード","決算発表日","次回決算発表日","配当CAGR"] #各列名
CSVCN1, CSVCN2, CSVCN3, CSVCN4, CSVCN5, CSVCN6 = range(len(CSV_COLNAME))
# JPX トップページ マーケット情報 統計情報（株式関連） その他統計資料 東証上場銘柄一覧
# 東証上場銘柄一覧 https://www.jpx.co.jp/markets/statistics-equities/misc/01.html
# から "コード" と "銘柄名" の列だけ残して IRCODE.csv として保存すれば使用可能
# (エラーが出るが、残りの列名が自動的に追加されるので、次からはエラーにならない)

# IRcache   = IRCache.CacheManager("irbank_cache.json") # 各種html情報の保存
# IRcache.stock_code    # 銘柄コードの指定
# IRcache.get( key )
# IRcache.set( key, value )
# IRcache.clear()

###########################################################################
import signal
from functools import wraps

def ignore_sigint(method):    # Ctrl+Cを無効化してmethodを呼び出す関数デコレーター
    @wraps(method)
    def wrapper(*args, **kwargs):
        print("(Ctrl+C 無効化)", end="")
        old = signal.signal(signal.SIGINT, signal.SIG_IGN)
        try:
            return method(*args, **kwargs)
        finally:
            signal.signal(signal.SIGINT, old)
            print("(Ctrl+C 有効)")
    return wrapper

###########################################################################
import csv

import datetime
import win32clipboard
import win32con

class CsvManager:
    def __init__(self, filename="IRCODE.csv", colnames=CSV_COLNAME):
        self.filename = filename
        self.colnames = colnames or []
        self.rows = self._load()

        self._stock_code = None
        self.row = None  # 現在の銘柄の行

    #def __del__(self):  #使わない
    #    self.update_row()           # キャッシュ保存
        
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc, tb):
        self.update_row()           # キャッシュ保存
        
    # --- ファイル読み込み ---
    def _load(self):
        # IRCODE.csv が存在しない場合は空の rows を返す
        if not os.path.exists(self.filename):
            print(f"(_load) {self.filename} が無いため、新規作成モードで開始します")
            return []

        rows = []
        with open(self.filename, newline='', encoding='CP932') as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append(row)
        return rows
    
    # --- ファイル保存 ---
    @ignore_sigint      # Ctrl+C を一時的に無視
    def save(self):
        print(f" {self.filename}保存 ", end="")
        with open(self.filename, "w", newline='', encoding='CP932') as f:
            writer = csv.DictWriter(f, fieldnames=self.colnames)
            writer.writeheader()
            writer.writerows(self.rows)

    # --- stock_code プロパティ ---
    @property
    def stock_code(self):
        return self._stock_code
    
    @stock_code.setter
    def stock_code(self, value):
        self._stock_code = value
        self.row = self._find_row() # キャッシュ読込
        if not self.row:
            self._stock_code = None

    # --- 値取得 ---
    def get(self, key):
        return self.row.get(key)
    # --- 値設定 ---
    def set(self, key, value):
        self.row[key] = value
        #self.update_row()

    # --- 行検索（code_colname は optional） ---
    def _find_row(self, code_colname=None):     # キャッシュ読込
        code_colname = code_colname or self.colnames[0]
        for r in self.rows:
            if r.get(code_colname) == self._stock_code:
                return r
        return None

    # --- 行更新（row プロパティを保存） ---
    def update_row(self):                       # キャッシュ保存
        if not self.row:
            return
        code_colname = self.colnames[0]
        for r in self.rows:
            if r.get(code_colname) == self._stock_code:
                r.update(self.row)
                self.save()
                return

    #def refresh_row(self):
    #    self.row = self._find_row()
        
    ###########################################################################
    def marge_stock_code(self, jpx_filename="jpx_stock_list.csv"):
    #jpx_stock_list.csv にあって IRCODE.csv に無い銘柄コードを
    #IRCODE.csv の末尾に追加して保存する

    # --- JPX の銘柄一覧を読み込む ---
        jpx_rows = []
        with open(jpx_filename, newline='', encoding='CP932') as f:
            reader = csv.DictReader(f)
            for row in reader:
                jpx_rows.append(row)

        code_col = self.colnames[CSVCN1]   # "コード"
        name_col = self.colnames[CSVCN2]   # "銘柄名"

        # --- 既存の IRCODE.csv のコード一覧 ---
        existing_codes = {r.get(code_col) for r in self.rows}

        added_count = 0

        # --- JPX の全銘柄をチェック ---
        for r in jpx_rows:
            code = r.get(code_col)
            name = r.get(name_col)

            if not code:
                continue

            # 既に IRCODE.csv にある → スキップ
            if code in existing_codes:
                continue

            # --- 新規追加 ---
            new_row = {
                code_col: code,
                name_col: name,
                self.colnames[2]: "",   # IRBANKコード
                self.colnames[3]: "",   # 決算発表日
                self.colnames[4]: ""    # 次回決算発表日
            }

            self.rows.append(new_row)
            added_count += 1

        # --- 保存 ---
        if added_count > 0:
            self.save()
            print(f"(marge_stock_code) {added_count} 件の銘柄を追加しました")
        else:
            print("(marge_stock_code) 追加する銘柄はありませんでした")

    ###########################################################################
    def copy_past_next_announcement_codes(self):
        # 次回決算発表日が過去の銘柄コードをクリップボードにコピーする
        # Excel に貼り付けると縦に並ぶ
        code_col = self.colnames[CSVCN1]      # "コード"
        next_col = self.colnames[CSVCN5]      # "次回決算発表日"

        today = datetime.date.today()
        result_codes = []

        for r in self.rows:
            next_dt = r.get(next_col)

            if not next_dt:
                continue

            try:
                dt = datetime.datetime.strptime(next_dt, "%Y-%m-%d").date()
            except:
                continue

            if dt < today:
                result_codes.append(r.get(code_col))

        # --- 改行区切りでクリップボードへコピー ---
        text = "\n".join(result_codes)

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        win32clipboard.CloseClipboard()

        print(f"(copy_past_next_announcement_codes) {len(result_codes)} 件コピーしました")



            
###########################################################################
import json
import os
import tempfile

class CacheManager:
    def __init__(self, filename="irbank_cache.json"):
        self.filename = filename
        self.cache = self._load()

        self._stock_code = None

    # --- stock_code プロパティ ---
    @property
    def stock_code(self):
        return self._stock_code

    @stock_code.setter
    def stock_code(self, value):
        self._stock_code = value

    # --- ファイル読み込み ---
    def _load(self):
        if os.path.exists(self.filename):
            with open(self.filename, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    # --- 保存 ---
    @ignore_sigint      # Ctrl+C を一時的に無視
    def save(self):
        dir_name = os.path.dirname(self.filename) or "."
        fd, tmp_path = tempfile.mkstemp(dir=dir_name)

        try:    # Ctrl+Cで jsonファイルが壊れないようにする
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(self.cache, f, ensure_ascii=False, indent=2)

            print(f" {self.filename}保存 ", end="")
            os.replace(tmp_path, self.filename)
        except Exception:
            os.remove(tmp_path)
            raise

        #with open(self.filename, "w", encoding="utf-8") as f:
        #    json.dump(self.cache, f, ensure_ascii=False, indent=2)

    def get(self, key):             # --- 値取得 ---
        if not self._stock_code:
            return None
        return self.cache.get(self._stock_code, {}).get(key)
    
    def set(self, key, value):      # --- 値設定 ---
        if not self._stock_code:
            return
        if self._stock_code not in self.cache:
            self.cache[self._stock_code] = {}
        self.cache[self._stock_code][key] = value

    def clear(self):                # --- クリア ---
        if not self._stock_code:
            return
        self.cache[self._stock_code] = {}
        #self.save()

    def rename_key(self, old_key, new_key): # key変更
        data = self.cache.get(self.stock_code, {})

        if old_key not in data:
            return False

        data[new_key] = data.pop(old_key)
        return True

            
        
#############################################################################################
# 実行 (テストドライバ)
if __name__ == "__main__":
    if 1 :                          # IRBANK内部コード取得
        with CsvManager() as Csv:
            Csv.marge_stock_code()

    elif 0 :
        IRcsv = CsvManager("IRCODE.csv", CSV_COLNAME)
        IRcsv.marge_stock_code("jpx_stock_list.csv")

    elif 0 :                          # 次回決算発表日が過去のものをコピー
        with CsvManager() as Csv:
            Csv.copy_past_next_announcement_codes()
        
