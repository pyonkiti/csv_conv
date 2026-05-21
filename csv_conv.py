# ExcelをCSVに出力する機能です。
# 送付書をExcelのWord差し込み印刷から楽楽販売に移行するために利用

import pandas as pd
import csv
import unicodedata
import configparser
import re
import sys
from pathlib import Path

file_path = r"C:\vagrant\csv_conv"                                      # 出力CSVパス
delimiter_type = "comma"                                                # "comma" または "tab"
encoding_type = "utf-8"                                                 # "utf-8" または "ansi"
max_rows = 150                                                          # 最大読込行数

# 共通処理
class PROC_HEAD:

    # INIファイルの読み込み
    def get_ini():
        try:
            retbln = False

            # INIファイルの読み込み
            config = configparser.ConfigParser()
            config.read(fr"{file_path}\\SYSTEM.INI", encoding = "utf-8-sig")
            
            excel_file =  config.get('EXCEL', 'EXCEL_FILE')             # Excelファイル名
            retbln = True

        except Exception as e:
            msg_err = "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            print(msg_err)
        finally:
            return retbln, excel_file

# Excel処理クラス
class PROC_EXCEL:

    # ExcelをCSVにコンバート
    def create_excel(sheet_names):
        
        try:
            retbln = False

            # INIファイルを読み込む
            ret, excel_file = PROC_HEAD.get_ini()
            if not ret: return retbln
                
            delimiter = "," if delimiter_type == "comma" else "\t"                          # 区切り文字設定
            encoding = "utf-8-sig" if encoding_type.lower() == "utf-8" else "cp932"         # エンコード設定

            Path(f"{file_path}\\csv").mkdir(parents = True, exist_ok = True)                # 出力フォルダ作成
            sheet_list = pd.ExcelFile(rf"{file_path}\{excel_file}").sheet_names             # シート一覧を取得
            
            match = re.search(r"【(.*?)】", excel_file)                                     # Excelファイルの中からファイル区分を取得
            if match:
                fil_kbn = match.group(1)
            else:
                fil_kbn= "_"

            # シートごとにCSV出力
            for index, sheet_name in enumerate(sheet_names):
                
                # シート存在チェック
                if not sheet_name in sheet_list: continue
                    
                # Excel読み込み
                df = pd.read_excel(rf"{file_path}\{excel_file}", sheet_name, header = None, dtype = str)
                
                # 開始行・開始列適用
                df = df.iloc[0:, 2:]

                # 最大行数制限
                if max_rows is not None:
                    df = df.iloc[:max_rows]

                # NaNを空文字へ変換
                df = df.fillna("")
                
                seikyu_col = 7                                                              # 請求書式数の列数
                meisai_col = 8                                                              # 明細式数の列数

                # 請求書式数を全角 → 半角変換
                df.iloc[:, seikyu_col] = df.iloc[:, seikyu_col].apply( lambda x: unicodedata.normalize("NFKC", str(x)) )

                # 明細式数を全角 → 半角変換
                df.iloc[:, meisai_col] = df.iloc[:, meisai_col].apply( lambda x: unicodedata.normalize("NFKC", str(x)) )
                
                # 固定値の追加
                df.insert(0, "納付区分", "明細ⅩⅩ")

                # ヘッダ名設定
                match len(df.columns):
                    case 10:
                        header_col = ["納付区分","郵便番号","住所1","住所2","会社名","部署名","担当名","請求名","請求書式数","明細式数"]
                    case 11:
                        header_col = ["納付区分","郵便番号","住所1","住所2","会社名","部署名","担当名","請求名","請求書式数","明細式数", "備考１"]
                    case 12:
                        header_col = ["納付区分","郵便番号","住所1","住所2","会社名","部署名","担当名","請求名","請求書式数","明細式数", "備考１", "備考２"]
                    case _:
                        header_col = ""

                df.columns = header_col
                            
                # 出力ファイル名
                output_file = Path(f"{file_path}\\csv") / f"{fil_kbn}_{sheet_name}.csv"

                # CSV出力
                df.to_csv(output_file, index = False, header = True, sep = delimiter, encoding = encoding, quoting = csv.QUOTE_ALL)

                print(f"{sheet_name}.csv を作成しました。")

            retbln = True

        except Exception as e:
            msg_err = "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            print(msg_err)
        finally:
            return retbln
        

# CSV処理クラス
class PROC_CSV:

    # CSVファイルにカラムを追加する
    def addcow_csv():

        # CSVフォルダ
        csv_dir = Path("./csv")

        # CSVファイルを取得
        csv_files = csv_dir.glob("*.csv")

        for csv_file in csv_files:

            # ファイル名から支払期間名を取得 （例: _ABC_202405.csv → ABC）
            match = re.match(r'^_(.*?)_.*\.csv$', csv_file.name)

            if not match:
                print(f"ファイル名形式が不正のためスキップ: {csv_file.name}")
                continue

            match match.group(1):
                case "毎月":         shiharai = "01"
                case "半期請求":     shiharai = "11"
                case "年度末請求":   shiharai = "22"
                case "4月.10月請求": shiharai = "13"
                case "2月.8月請求":  shiharai = "12"
                case _:             shiharai = ""

            # 出力ファイル名
            output_file = csv_file.with_name(f"{csv_file.stem}_new.csv")

            # CSV読み込み・書き込み
            with open(csv_file, mode="r", encoding="utf-8-sig", newline="") as infile, \
                open(output_file, mode="w", encoding="utf-8-sig", newline="") as outfile:

                reader = csv.reader(infile)

                # QUOTE_ALL により全項目を "" で囲む
                writer = csv.writer(outfile, quoting=csv.QUOTE_ALL)

                for row_index, row in enumerate(reader):

                    # 1カラム目と2カラム目の間に追加
                    if row_index == 0:
                        new_row = [ row[0], "支払先期間コード", *row[1:] ]        # ヘッダ行
                    else:
                        new_row = [ row[0], shiharai, *row[1:] ]               # データ行

                    writer.writerow(new_row)

            print(f"出力完了: {output_file}")


# ----------------------------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------------------------
if __name__ == "__main__":

    match sys.argv[1]:
        case "1":
            
            target_sheets = ["差し込み可能 (明細なし)", "差し込み可能", "WA専用請求書なし", "早めの送付先", "差し込み可能（2月のみ）", "差し込み可能（2月・8月）", "差し込み可能（4月・10月）"]    # 出力対象シート
    
            # ExcelをCSVにコンバート
            ret = PROC_EXCEL.create_excel(target_sheets)
            if ret:
                print("CSV出力完了")
            else:
                print("処理が異常終了しました")

        case "2":

            # CSVファイルにカラム（支払期間名）を追加
            ret = PROC_CSV.addcow_csv()
            
        case _:
            None
