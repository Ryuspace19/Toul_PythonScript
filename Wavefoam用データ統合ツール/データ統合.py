import pandas as pd
from tkinter import Tk, filedialog
import os
import re

# ファイル選択ダイアログを表示する関数
def select_files_dialog(title, filetypes):
    root = Tk()
    root.withdraw()  # GUIウィンドウを非表示にする
    file_paths = filedialog.askopenfilenames(title=title, filetypes=filetypes)
    root.destroy()
    return file_paths

# 保存先ファイルダイアログを表示する関数
def save_file_dialog(title, defaultextension, filetypes):
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(title=title, defaultextension=defaultextension, filetypes=filetypes)
    root.destroy()
    return file_path

# 数値順でファイルをソートする関数
def numeric_sort(file_list):
    return sorted(file_list, key=lambda x: int(re.search(r'\d+', os.path.basename(x)).group()))

# ヘッダーを作成する関数
def create_header(worksheet, vrd_values):
    col_idx = 1
    for vrd in vrd_values:
        worksheet.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx + 3)
        worksheet.cell(row=1, column=col_idx).value = f"Vrd : {vrd}V"
        col_idx += 4
    columns = ["Id (A)", "Vgs (V)", "Vrd (V)", "Id (mA)"] * len(vrd_values)
    for col_num, column_title in enumerate(columns, 1):
        worksheet.cell(row=2, column=col_num).value = column_title

    # Y列のヘッダー
    worksheet.cell(row=2, column=col_num + 1).value = "File Name"

# メイン処理
def merge_excel_files():
    # 入力ファイルの選択
    input_files = list(select_files_dialog("Excelファイルを選択してください（複数選択可能）", [("Excel Files", "*.xlsx")]))
    if not input_files:
        print("ファイルが選択されませんでした。")
        return

    # 数値順にソート
    input_files = numeric_sort(input_files)

    # データ結合用のリスト
    combined_data = []

    # Vrdの範囲
    vrd_values = ["0", "1", "2", "3", "4", "5"]

    # 各ファイルのデータを読み込み、1行空けて結合
    for file in input_files:
        data = pd.read_excel(file, header=1)  # ヘッダーは2行目から
        data["File Name"] = os.path.basename(file)  # ファイル名を新しい列として追加
        combined_data.append(data)
        # 空白行を挿入
        combined_data.append(pd.DataFrame([[]]))

    # 結合データフレームの作成
    merged_data = pd.concat(combined_data, ignore_index=True)

    # 保存先ファイルの選択
    output_file = save_file_dialog("結合結果を保存するファイルを指定してください", ".xlsx", [("Excel Files", "*.xlsx")])
    if not output_file:
        print("保存先が選択されませんでした。")
        return

    # Excelライターを使用してファイルに書き出し
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        merged_data.to_excel(writer, index=False, startrow=2, header=False)
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        # ヘッダーを作成
        create_header(worksheet, vrd_values)

    print(f"結合結果が正常に保存されました。保存先: {output_file}")

# 実行
merge_excel_files()
