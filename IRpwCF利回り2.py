
import IRpwBB5
import pandas as pd             # py -m pip install pandas


#############################################################################################
def makeCFYield(dfCF, dfV, dfD):     #CF推移と価値算定を受けて、CF利回り一覧を返す
    if dfCF.empty or dfV.empty:
        print("\n####CF推移or価値算定のデータなし####\n\n")
        return pd.DataFrame()
    
    df = dfCF[['単位', '営業CF']].join(dfV[['時価総額']], how='outer') #欠損も残す
    #df = df.replace('-', pd.NA).dropna()    # '-' や NaN を含む行を削除

    # 数値化（営業CFはカンマ除去 → int）百万円単位
    #df['営業CF'] = df['営業CF'].str.replace(',', '').astype(int) * df['単位'].astype(int)
    df["営業CF"] = (
        pd.to_numeric(df["営業CF"].str.replace(",", ""), errors="coerce")
        * pd.to_numeric(df["単位"], errors="coerce")
    )
    # 時価総額データなし 8698 マネックスグループ 8473 ＳＢＩホールディングス 366A

    
    # 時価総額を百万円単位で数値化（兆・億に対応）
    def parse_marketcap(x):
        if isinstance(x, str):
            if '百万' in x:   # 学びエイド（184A）「915百万」
                x = x.replace('百万', '')
                print(f"(makeCFYield){x}", type(x))
                return float(x)
            if '万' in x:
                x = x.replace('万', '').replace('億', '')
                return float(x) * 1_0000 / 100_0000
                # テクセンドフォトマスク（429A）時価総額「3729億4639万」
            if '億' in x:
                x = x.replace('億', '').replace('兆', '')
                return float(x) * 1_0000_0000 / 100_0000
                # 
            if '兆' in x:
                x = x.replace('兆', '').replace('京', '')
                return float(x) * 1_0000_0000_0000 / 100_0000
        try:
            return float(x)
        except:
            print(f"(makeCFYield){x}", type(x))
            return pd.NA

    #df['時価総額'] = df['時価総額'].apply(parse_marketcap).astype(float)
    df['時価総額'] = pd.to_numeric(df['時価総額'].apply(parse_marketcap), errors="coerce")
        
    df['CF利回り'] = df['営業CF'] / df['時価総額'] #CF利回り（営業CF ÷ 時価総額）
    df.drop(columns=['単位'], inplace=True)       #単位列を削除

    if "配当利回り" in dfD.columns and not dfV.empty:
        df = df.join( dfD[['配当利回り']], how='outer')
    return df.tail(15)

#############################################################################################
# 実行 (テストドライバ)
if __name__ == "__main__":
    import pyperclip                #pip install pyperclip
    import sys
    import time         #sleep
    import pythoncom    #pythoncom.PumpWaitingMessages()
    from datetime import datetime, time as dtime

    import win32com.client as win32

    win32.gencache.EnsureDispatch("Excel.Application")  # 型ライブラリをロード
    # AttributeError: module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x8'
    # has no attribute 'CLSIDToClassMap'
    # ⇒ キャッシュ破損: %LOCALAPPDATA%\Temp\gen_py フォルダを削除
    from win32com.client import constants   # xlRight
    
    excel = win32.Dispatch("Excel.Application")
    ws = excel.ActiveSheet
    
    adj1, adj2 = 1, 1
    #自セルを参照する Offset(adj1, adj2) を書く際のパラメータ
    #本来(VBA等)は 0, 0だが、1, 1でないと目論見通りにならないので明示する

    def is_in_time_range():         #14:30～16:30の間なら真 (ループしない)
        now = datetime.now().time()
        return dtime(14, 30) <= now <= dtime(16, 30)

    def PasteGraph( off1, off2 ):   # off1行下、off2列右にPasteしてサイズ調整
        success = False
        for _ in range(10):
            try:
                ws.Paste(Destination=cActv.Offset( off1 + adj1, off2 + adj2))
                success = True
                break
            except Exception:
                time.sleep(0.1)
        if not success:
            raise Exception("貼り付け失敗")
        
        shp = ws.Shapes(ws.Shapes.Count)
        shp.ScaleHeight(0.5, True)   # 高さ 50%
        shp.ScaleWidth(0.5, True)    # 幅 50%
        shp.Left = shp.Left + 10
        shp.Top = shp.Top + 5
        time.sleep(0.1)

    def safe_set_value(cell, value, retries=10):
        for i in range(retries):
            try:
                cell.Value = value
                return
            except Exception as e:
                time.sleep(0.1)
        raise e
    

    with IRpwBB5.IRBankBrowser() as irBB:

        irBB.ref_list = ["cf", "value", "dividend", "per"] # 並行読み込み
        
        cActv = excel.ActiveCell
        while cActv.Text:
            cNext = cActv.Offset(1 + adj1, adj2)    #次のセル

            #s = pyperclip.paste()  # クリップボードの内容を取得
            #s = s.replace("\r", "").replace("\n", "")
            s = excel.ActiveCell.Text
            irBB.stock_code = s                 #銘柄コード
            
            if not irBB.stock_code: #銘柄コードでないものはパス
            #if not irBB.stock_code or not irBB.cache_cleared: #新決算情報が無い場合もパス
                cNext.Select()
                cActv = excel.ActiveCell
                time.sleep(1)
                continue

            # URLの貼り付け
            #cActv.Offset( adj1, adj2 + 1).Value = irBB.stock_name
            safe_set_value(cActv.Offset( adj1, adj2 + 1), irBB.stock_name)
            urls = ((irBB.url_cf, 2), (irBB.url_dividend, 5), (irBB.url_value, 3),
                    (irBB.url_per, 8), (irBB.url_pl, 11))
            for url,col in urls:
                #cActv.Offset( adj1, adj2 + col).Value = url
                safe_set_value(cActv.Offset( adj1, adj2 + col), url)
            time.sleep(0.001)
            pythoncom.PumpWaitingMessages()     #DoEvents

            # CF利回り一覧表の貼り付け
            dfCF = irBB.read_cf_table()           # CF推移テーブル(/cf)を取得
            dfV  = irBB.read_value_table()        # 価値算定テーブル(/value)読込
            dfD  = irBB.read_dividend_table()     # 配当金の推移テーブル(/dividend)読込
            
            df = makeCFYield( dfCF, dfV, dfD )
            tsv = df.to_csv(sep="\t")
            tsv = tsv.replace("\r\n", "\n").replace("\r", "\n") # 改行を\n に統一
            print(tsv, end="")
            #pyperclip.copy(tsv) #クリップボードにコピー Paste直前に移動

            row = cActv.Row             # アクティブセルの行番号を取得
            i = df.shape[0] + 3         # dfの行数 + 3 (タイトル1行、余白2行)
            i = 19 if i < 19 else i                 # 最小19行
            ws.Rows(f"{row+1}:{row+i}").Insert()    # 下に挿入
        
            pyperclip.copy(tsv) #クリップボードにコピー
            time.sleep(0.1)
            pythoncom.PumpWaitingMessages()     #DoEvents
            ws.Paste(Destination=cActv.Offset( adj1 + 1, adj2 + 1)) #貼り付け
            
            time.sleep(0.1)
            cStart = cActv.Offset( adj1 + 1, adj2 + 4)
            cEnd   = cActv.Offset( adj1 + i, adj2 + 4)
            rng = excel.ActiveSheet.Range(cStart, cEnd) #CF利回り列
            rng.NumberFormat = "0.00%"          # 書式 パーセント 小数2桁

            # 次回決算発表日, 直近決算発表日, 配当金CAGRの貼り付け
            cStart = cActv.Offset( adj1 + 1, adj2)
            cEnd   = cActv.Offset( adj1 + 3, adj2)
            rng = excel.ActiveSheet.Range(cStart, cEnd)
            rng.Value = [[irBB.next_announcement],[irBB.announcement],[irBB.Dividend_CAGR]]
            cActv.Offset( adj1 + 1, adj2).NumberFormat = "M/D"
            cActv.Offset( adj1 + 3, adj2).NumberFormat = "0.00%" 

            # URLにハイパーリンクを付ける
            for url,col in urls:
                cActv.Offset( adj1, adj2 + col).HorizontalAlignment = constants.xlRight
                ws.Hyperlinks.Add(Anchor=cActv.Offset( adj1, adj2 + col),
                                  Address=url)

            #各種グラフ貼り付け
            irBB.load_pl_page()         # 会社業績ページ(/pl)読込
            if irBB.copy_net_income_graph():    # 当期純利益グラフ（SVG #2）
                PasteGraph(1, 10)
        
            if irBB.copy_sales_graph():     # 売上高グラフ（SVG #0）
                PasteGraph(10, 8)

            irBB.load_per_page()        # PER推移ページ(/per)読込
            if irBB.copy_per_graph():       # PERグラフ
                PasteGraph(1, 6)
           
            #次の銘柄へ移動
            cNext.Select()
            cActv = excel.ActiveCell

            if is_in_time_range():      # 14:30-16:30の間はループしない
            #if 0:
                print("\nループ処理中断\n")
                while cActv.Text:       # 最後まで行く
                    cNext = cActv.Offset(1 + adj1, adj2)    #次のセル
                    cNext.Select()
                    cActv = excel.ActiveCell
                continue

            # ループ中断するならこのタイミングで
            # Pythonメッセージ出力窓に Ctrl+C する
            irBB.stock_code = None
            for i in range(30):
                print("-", end="")
                time.sleep(0.1)
                pythoncom.PumpWaitingMessages()     #DoEvents
            print("---------")

    
    
