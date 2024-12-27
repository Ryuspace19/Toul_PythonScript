#コマンドプロンプトで実行する場合は以下の順に実行すること
#pip install PyPDF2
#pip install tkinter
#cd このスクリプトがあるパス
#python rotate_pdf.py

import os
from PyPDF2 import PdfReader, PdfWriter
from tkinter import Tk, filedialog

def rotate_even_pages(input_pdf, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        if (i + 1) % 2 == 0:  # 偶数ページの場合
            page.rotate(180)  # ページを180度回転
        writer.add_page(page)

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

def main():
    # GUIを表示しない設定
    Tk().withdraw()

    # 入力ファイルの選択
    input_pdf = filedialog.askopenfilename(
        title="偶数ページを回転するPDFを選択してください",
        filetypes=[("PDFファイル", "*.pdf")]
    )

    if not input_pdf:  # ファイルが選択されなかった場合
        print("PDFファイルが選択されませんでした。")
        return

    # 保存先フォルダの選択
    output_folder = filedialog.askdirectory(title="変換後のPDFを保存するフォルダを選択してください")

    if not output_folder:  # フォルダが選択されなかった場合
        print("保存先フォルダが選択されませんでした。")
        return

    # 出力ファイルの名前を生成
    output_filename = os.path.splitext(os.path.basename(input_pdf))[0] + "_回転済み.pdf"
    output_pdf = os.path.join(output_folder, output_filename)

    # 偶数ページを回転
    rotate_even_pages(input_pdf, output_pdf)

    print(f"変換が完了しました。保存先: {output_pdf}")

if __name__ == "__main__":
    main()
