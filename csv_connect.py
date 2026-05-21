# 複数のCSVファイルを１つのCSVファイルに結合する。
# 各CSVファイルの件数を調べる。

import pandas as pd
from pathlib import Path
import csv
import sys

csv_folder = Path("./csv")                      # CSVフォルダ
output_file = "merged.csv"                      # 出力ファイル

# CSVファイル
class CSV:
    # CSVファイルを結合する
    def csv_connect():

        # CSVファイル一覧取得（merged.csv は除外）
        csv_files = [ file for file in csv_folder.glob("*.csv") if file.name != f"{output_file}" ]
            
        # DataFrame格納用
        df_list = []

        # CSV読込
        for file in csv_files:

            print(f"読込: {file.name}")

            # CSVファイルを読み込む
            df = pd.read_csv( file, dtype = str, encoding = "utf-8" )

            # カラム名を統一（「備考１」「備考２」→「備考」）
            df.rename( columns = { "備考１": "備考", "備考２": "備考" }, inplace = True )

            # 同名カラムが複数できた場合は1つにまとめる
            if list(df.columns).count("備考") > 1:

                # 重複した「備考」カラムを横方向に結合
                remarks = df.loc[:, df.columns == "備考"]

                # 空でない値を優先して1列に統合
                df["備考"] = remarks.bfill(axis=1).iloc[:, 0]

                # 重複カラム削除
                df = df.loc[:, ~df.columns.duplicated()]

            df_list.append(df)

        # CSV結合
        merged_df = pd.concat(df_list, ignore_index = True)

        # CSV出力
        merged_df.to_csv( csv_folder / output_file, index = False, encoding = "utf-8", quoting = csv.QUOTE_ALL)
        
        print(f"完了: {csv_folder}/{output_file}")

    # CSVファイルのデータ件数を表示
    def csv_count():

        # 除外対象のCSVファイル名
        exclude_files = [f"{output_file}","test.csv"]

        # 総合計件数
        total_count = 0

        # csvファイルをループ
        for csv_file in csv_folder.glob("*.csv"):

            # 除外対象ならスキップ
            if csv_file.name in exclude_files:
                continue

            # データ件数
            row_count = 0

            # CSV読み込み
            with open(csv_file, mode = "r", encoding = "utf-8", newline = "") as f:
                reader = csv.reader(f)

                # ヘッダ行をスキップ
                next(reader, None)

                # データ行をカウント
                for row in reader:
                    row_count += 1

            # 件数表示
            print(f"{row_count:>3}件 - {csv_file.name}")

            # 総合計に加算
            total_count += row_count

        # 総合計表示
        print("--------------------")
        print(f"{total_count:>3}件 - 総合計")


match sys.argv[1]:
    case "1":
        CSV.csv_connect()                       # CSVファイルを結合する
    case "2":
        CSV.csv_count()                         # CSVファイルのデータ件数を表示
    case _:
        None
