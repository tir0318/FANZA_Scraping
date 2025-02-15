#FANZAの画像スクレイピング　現在

import os
import requests
from bs4 import BeautifulSoup
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def download_images(url, excel_path):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"}
    cookie = {'age_check_done': '1'}
    session = requests.session()
    session.get(url)

    # Excelファイルからフォルダ名を取得
    wb = openpyxl.load_workbook(excel_path)
    sheet = wb.active  # アクティブシートを取得
    folder_name = sheet["C2"].value  # C2セルの値を取得
    folder_directory = os.path.join(r"", folder_name) #フォルダのパスを入力

    # フォルダ作成
    os.makedirs(folder_directory, exist_ok=True)
    for i in range(1, 21):  # 1から20の番号でサブフォルダを作成
        subfolder_name = str(i).zfill(2)  # "01", "02", ..., "20" の形式
        os.makedirs(os.path.join(folder_directory, subfolder_name), exist_ok=True)

    with open(r"", 'r', encoding='UTF-8') as f: #URLリストファイルのパスを入力
        lines = f.readlines()
        
        for idx, line in enumerate(reversed(lines), start=1):
            soup = BeautifulSoup(session.get(line.strip(), headers=headers, cookies=cookie).content, 'html.parser')
            package_image = soup.find('div', class_='center').find('a').get("href")
            package_image_name = package_image.split('/')[-1]
            print(f"Downloading package image: {package_image}")

            # フォルダ番号を2桁の形式に変換
            subfolder = os.path.join(folder_directory, str((idx - 1) % 20 + 1).zfill(2))  
            r = requests.get(package_image)
            with open(os.path.join(subfolder, package_image_name), 'wb') as img_file:
                img_file.write(r.content)

            try:
                sample_image_list = soup.find('div', class_='d-zoomimg-sm').find_all('a')
                for sample_image in sample_image_list:
                    sample_image_url = sample_image.get("href")
                    sample_image_name = sample_image_url.split('/')[-1]
                    r = requests.get(sample_image_url)
                    with open(os.path.join(subfolder, sample_image_name), 'wb') as img_file:
                        img_file.write(r.content)
            except Exception as e:
                print(f"Error downloading sample images for URL {line.strip()}: {e}")

def select_excel_and_run(root):
    # ファイル選択ダイアログを開く
    file_path = filedialog.askopenfilename(
        title="Excelファイルを選択してください",
        filetypes=[("Excel Files", "*.xlsx;*.xls")],
        initialdir=os.path.expanduser("~/Documents")  # 初期ディレクトリをDownloadsに設定
    )

    if not file_path:
        messagebox.showwarning("警告", "ファイルが選択されていません")
        return

    # 確認メッセージ
    if not messagebox.askyesno("確認", f"選択したファイル: {file_path}\nこのファイルで処理を開始しますか？"):
        return

    # GUIを閉じる
    root.destroy()

    try:
        # ダウンロード処理を実行
        url_certification = 'https://www.dmm.co.jp/age_check/=/declared=yes/?rurl=https%3A%2F%2Fwww.dmm.co.jp%2Ftop%2F'
        download_images(url_certification, file_path)

        # 処理終了メッセージを表示
        messagebox.showinfo("完了", "ダウンロードが完了しました")

        # プログラムを終了
        sys.exit()
    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
        sys.exit()

def main():
    # Tkinterウィンドウを設定
    root = tk.Tk()
    root.title("Excelファイル選択")

    # ウィンドウを常に最前面に設定
    root.attributes("-topmost", True)

    # ウィンドウのサイズを設定
    root.geometry("400x200")

    # 説明ラベル
    label = tk.Label(root, text="処理に使用するExcelファイルを選択してください", font=("Arial", 12))
    label.pack(pady=20)

    # 実行ボタン
    button = tk.Button(root, text="Excelファイルを選択", command=lambda: select_excel_and_run(root), font=("Arial", 12), bg="lightblue")
    button.pack(pady=20)

    # Tkinterメインループ開始
    root.mainloop()

if __name__ == "__main__":
    main()
