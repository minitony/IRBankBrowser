
import IRpwBB5
import pandas as pd             # py -m pip install pandas


#############################################################################################
def makeCFYield(dfCF, dfV, dfD):     #CF推移と価値算定を受けて、CF利回り一覧を返す
    if dfCF.empty or dfV.empty:
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
                # テクセンドフォトマスク（429A）時価総額「3729億4639万」
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

    if "配当利回り" in dfD.columns:
        df = df.join( dfD[['配当利回り']], how='outer')
    return df.tail(15)

#############################################################################################
# 実行 (テストドライバ)
if __name__ == "__main__":
    import pyperclip                #pip install pyperclip
    import sys
    import time         #sleep
    import pythoncom    #pythoncom.PumpWaitingMessages()

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

    def PasteGraph( off1, off2 ):   # off1行下、off2列右にPasteしてサイズ調整
        ws.Paste(Destination=cActv.Offset( off1 + adj1, off2 + adj2))
        shp = ws.Shapes(ws.Shapes.Count)
        shp.ScaleHeight(0.5, True)   # 高さ 50%
        shp.ScaleWidth(0.5, True)    # 幅 50%
        shp.Left = shp.Left + 10
        shp.Top = shp.Top + 5
    

    with IRpwBB5.IRBankBrowser() as irBB:

        irBB.ref_list = ["cf", "value", "dividend", "per"] # 並行読み込み
        
        cActv = excel.ActiveCell
        while cActv.Text:
            cNext = cActv.Offset(1 + adj1, adj2)    #次のセル

            #s = pyperclip.paste()  # クリップボードの内容を取得
            #s = s.replace("\r", "").replace("\n", "")
            s = excel.ActiveCell.Text
            irBB.stock_code = s                 #銘柄コード
            pythoncom.PumpWaitingMessages()     #DoEvents
            time.sleep(0.001)
            
            if not irBB.stock_code: #銘柄コードでないものはパス
            #if not irBB.stock_code or not irBB.cache_cleared: #新決算情報が無い場合もパス
                cNext.Select()
                cActv = excel.ActiveCell
                time.sleep(1)
                continue

            # URLの貼り付け
            cActv.Offset( adj1, adj2 + 1).Value = irBB.stock_name
            urls = ((irBB.url_cf, 2), (irBB.url_dividend, 5), (irBB.url_value, 3),
                    (irBB.url_per, 8), (irBB.url_pl, 11))
            for url,col in urls:
                cActv.Offset( adj1, adj2 + col).Value = url
            pythoncom.PumpWaitingMessages()     #DoEvents
            time.sleep(0.001)

            # CF利回り一覧表の貼り付け
            dfCF = irBB.read_cf_table()           # CF推移テーブル(/cf)を取得
            dfV  = irBB.read_value_table()        # 価値算定テーブル(/value)読込
            dfD  = irBB.read_dividend_table()     # 配当金の推移テーブル(/dividend)読込
        
            df = makeCFYield( dfCF, dfV, dfD )
            tsv = df.to_csv(sep="\t")
            # 改行コードを強制的に \n に統一
            tsv = tsv.replace("\r\n", "\n").replace("\r", "\n")
            print(tsv, end="")
            pyperclip.copy(tsv) #クリップボードにコピー

            row = cActv.Row              # アクティブセルの行番号を取得
            i = df.shape[0] + 3         # dfの行数 + 3 (タイトル1行、余白2行)
            i = 19 if i < 19 else i     # 19行挿入
            ws.Rows(f"{row+1}:{row+i}").Insert()    # 下に i行挿入
        
            ws.Paste(Destination=cActv.Offset( adj1 + 1, adj2 + 1))
            cStart = cActv.Offset( adj1 + 1, adj2 + 4)
            cEnd   = cActv.Offset( adj1 + i, adj2 + 4)
            rng = excel.ActiveSheet.Range(cStart, cEnd)
            rng.NumberFormat = "0.00%"          # 書式 パーセント 小数2桁

            # 次回決算発表日と直近決算発表日の貼り付け
            cActv.Offset( adj1 + 1, adj2).Value = irBB.next_announcement
            cActv.Offset( adj1 + 1, adj2).NumberFormat = "M/D"
            cActv.Offset( adj1 + 2, adj2).Value = irBB.announcement
            # 配当金CAGRの貼り付け
            cActv.Offset( adj1 + 3, adj2).Value = irBB.Dividend_CAGR
            cActv.Offset( adj1 + 3, adj2).NumberFormat = "0.00%" 
            pythoncom.PumpWaitingMessages()     #DoEvents
            time.sleep(0.001)

            # URLにハイパーリンクを付ける
            for url,col in urls:
                cActv.Offset( adj1, adj2 + col).HorizontalAlignment = constants.xlRight
                ws.Hyperlinks.Add(Anchor=cActv.Offset( adj1, adj2 + col),
                                  Address=url)

            irBB.load_pl_page()         # 会社業績ページ(/pl)読込
            if irBB.copy_net_income_graph():    # 当期純利益グラフ（SVG #2）
                PasteGraph(1, 10)
        
            if irBB.copy_sales_graph():     # 売上高グラフ（SVG #0）
                PasteGraph(10, 8)
            pythoncom.PumpWaitingMessages()     #DoEvents
            time.sleep(0.001)

            irBB.load_per_page()        # PER推移ページ(/per)読込
            if irBB.copy_per_graph():       # PERグラフ
                PasteGraph(1, 6)
            pythoncom.PumpWaitingMessages()     #DoEvents
            time.sleep(0.001)

            cNext.Select()
            cActv = excel.ActiveCell

            for i in range(30):
                print("-", end="")
                pythoncom.PumpWaitingMessages()     #DoEvents
                time.sleep(0.1)
            print("---------")

    
    
