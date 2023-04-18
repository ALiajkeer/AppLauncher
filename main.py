import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import sqlite3
import os

APP_DEF_WIDTH = 300
APP_DEF_HEIGHT = 100


class DragAndDrop(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # DB接続
        self.conn = sqlite3.connect('app.db')
        self.cur = self.conn.cursor()
        # テーブル作成
        self.cur.execute('''CREATE TABLE IF NOT EXISTS apps
                           (id INTEGER PRIMARY KEY,
                            name TEXT,
                            path TEXT,
                            icon_path TEXT)''')
        self.conn.commit()

        # ウィンドウサイズ、タイトルの設定
        self.geometry(f'{APP_DEF_WIDTH}x{APP_DEF_HEIGHT}')
        self.minsize(APP_DEF_WIDTH, APP_DEF_HEIGHT)
        self.maxsize(APP_DEF_WIDTH+100, APP_DEF_HEIGHT+100)
        self.title(f'アプリランチャー')

        # フレーム
        self.frame_drag_drop = tk.LabelFrame(self)
        self.frame_drag_drop.textbox = tk.Text(self.frame_drag_drop)

        # テキストボックスに表示
        self.disp_app_info()

        # ドラッグアンドドロップ
        self.frame_drag_drop.textbox.drop_target_register(DND_FILES)
        self.frame_drag_drop.textbox.dnd_bind('<<Drop>>', self.func_drag_and_drop)

        # スクロールバー設定
        self.frame_drag_drop.scrollbar = tk.Scrollbar(self.frame_drag_drop, orient=tk.VERTICAL, command=self.frame_drag_drop.textbox.yview)
        self.frame_drag_drop.textbox['yscrollcommand'] = self.frame_drag_drop.scrollbar.set

        # 配置
        self.frame_drag_drop.textbox.grid(column=0, row=0, sticky=(tk.E, tk.W, tk.S, tk.N))
        self.frame_drag_drop.scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.frame_drag_drop.columnconfigure(0, weight=1)
        self.frame_drag_drop.rowconfigure(0, weight=1)
        self.frame_drag_drop.grid(column=0, row=0, padx=5, pady=5, sticky=(tk.E, tk.W, tk.S, tk.N))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

    # DBへ保存
    def save_app_info(self, name, path, icon_path):
        # データベースにアプリ情報を保存
        self.cur.execute("INSERT INTO apps(name, path, icon_path) VALUES (?, ?, ?)", (name, path, icon_path))
        self.conn.commit()

    # ドラッグアンドドロップ時、ファイルのパスを取得する
    def func_drag_and_drop(self, event):
        # ドロップされたファイルからアプリ情報を取得
        file_path = event.data.strip()
        # 拡張子なしのファイル名を抽出
        name = os.path.splitext(os.path.basename(file_path))
        full_path = event.data.strip('{}\'')
        icon_path = ''  # アイコンファイルパスはとりあえず空にしておく

        # アプリ情報をデータベースに保存
        self.save_app_info(name[0], full_path, icon_path)

        # テキストボックスに表示
        self.disp_app_info()

    # テキストボックスに表示
    def disp_app_info(self):
        # テキストボックスをクリア
        self.frame_drag_drop.textbox.configure(state='normal')
        self.frame_drag_drop.textbox.delete('1.0', tk.END)
        
        # テキストボックスにDBから読み込んだパスを表示する。DBが作成されていなければデフォルト値を表示する
        # DBからアプリ情報を読み込む
        self.cur.execute('SELECT name, path FROM apps ORDER BY id ASC')
        apps = self.cur.fetchall()
        if not apps:
            self.frame_drag_drop.textbox.insert(tk.END, "ここにファイルをドロップ")
        else:
            for app in apps:
                self.frame_drag_drop.textbox.insert(tk.END, f'{app[0]}\n')
        self.frame_drag_drop.textbox.configure(state='disabled')
        self.frame_drag_drop.textbox.see(tk.END)

    # アプリ終了時、DBを切断
    def __del__(self):
        self.conn.close()


def main():
    app = DragAndDrop()
    app.mainloop()


if __name__ == "__main__":
    main()
