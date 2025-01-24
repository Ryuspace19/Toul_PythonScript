import pandas as pd
import os
from tkinter import Tk, filedialog

# CSVデータをExcel形式に変換する関数
def convert_csv_to_excel(csv_file, output_file):
    # CSVファイルの読み込み
    csv_data = pd.read_csv(csv_file, skiprows=9)

    # Vrdの値を抽出し、データをフィルタリング
    vrd_values = csv_data['Vrd (V)'].unique()

    # 新しいデータフレームの作成
    formatted_data = pd.DataFrame()

    # 各Vrdごとにデータを整形して結合
    headers = []
    for vrd in vrd_values:
        subset = csv_data[csv_data['Vrd (V)'] == vrd]
        subset['Id (mA)'] = subset['Id (A)'] * 1000  # IdをmAに変換
        subset = subset.rename(columns={
            'Id (A)': 'Id (A)',
            'Vgs (V)': 'Vgs (V)',
            'Vrd (V)': 'Vrd (V)',
            'Id (mA)': 'Id (mA)'
        })
        # 列を整形して連結
        subset = subset[['Id (A)', 'Vgs (V)', 'Vrd (V)', 'Id (mA)']]
        subset.columns = ['Id (A)', 'Vgs (V)', 'Vrd (V)', 'Id (mA)']
        headers.extend([f'Vrd : {vrd}V', None, None, None])
        formatted_data = pd.concat([formatted_data, subset.reset_index(drop=True)], axis=1)

    # Excelライターを使用してヘッダーを追加
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        formatted_data.to_excel(writer, index=False, startrow=1, header=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # 1行目にセル結合とヘッダー記入
        col_idx = 1
        for vrd in vrd_values:
            worksheet.merge_cells(start_row=1, start_column=col_idx, end_row=1, end_column=col_idx+3)
            worksheet.cell(row=1, column=col_idx).value = f'Vrd : {vrd}V'
            col_idx += 4

        # 2行目に列名を追加
        for col_num, column_title in enumerate(formatted_data.columns, 1):
            worksheet.cell(row=2, column=col_num).value = column_title

# ファイル選択ダイアログを表示する関数
def select_files_dialog(title, filetypes):
    root = Tk()
    root.withdraw()  # GUIウィンドウを非表示にする
    file_paths = filedialog.askopenfilenames(title=title, filetypes=filetypes)
    root.destroy()
    return file_paths

# 保存先フォルダ選択ダイアログを表示する関数
def select_folder_dialog(title):
    root = Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_path

# 入力ファイルの選択
input_csv_files = select_files_dialog("CSVファイルを選択してください（複数選択可能）", [("CSV Files", "*.csv")])
if not input_csv_files:
    print("CSVファイルが選択されませんでした。")
    exit()

# 出力先フォルダの選択
output_folder = select_folder_dialog("出力先フォルダを選択してください")
if not output_folder:
    print("出力先フォルダが選択されませんでした。")
    exit()

# 一括変換の実行
for csv_file in input_csv_files:
    file_name = os.path.splitext(os.path.basename(csv_file))[0]  # ファイル名を取得
    output_file = os.path.join(output_folder, f"{file_name}.xlsx")
    convert_csv_to_excel(csv_file, output_file)
    print(f"{csv_file} を変換して {output_file} に保存しました。")

print("すべてのファイルの変換が完了しました！")
