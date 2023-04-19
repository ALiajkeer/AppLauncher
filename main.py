import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import sqlite3
import os

APP_DEF_WIDTH = 300
APP_DEF_HEIGHT = 100


class DragAndDrop(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # DBに接続し、テーブルを作成。既にテーブルが存在するなら作成しない。
        self.conn = sqlite3.connect('app.db')
        self.cur = self.conn.cursor()
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
        self.frame_drag_drop.grid(row=0, sticky="we")

        # リストボックス
        self.frame_drag_drop.listbox = tk.Listbox(self.frame_drag_drop, height=5, listvariable=tk.StringVar(), selectmode="single")
        self.frame_drag_drop.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # スクロールバー
        self.frame_drag_drop.scrollbar = tk.Scrollbar(self.frame_drag_drop, orient=tk.VERTICAL, command=self.frame_drag_drop.listbox.yview)
        self.frame_drag_drop.scrollbar.pack(side=tk.LEFT, fill=tk.Y)

        # リストボックスとスクロールバーを連動させる
        self.frame_drag_drop.listbox['yscrollcommand'] = self.frame_drag_drop.scrollbar.set

        # DBに登録しているアプリ情報を表示
        self.disp_app_info()

        # ドラッグアンドドロップ
        self.frame_drag_drop.listbox.drop_target_register(DND_FILES)
        self.frame_drag_drop.listbox.dnd_bind('<<Drop>>', self.func_drag_and_drop)

        # 配置
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

    # DBに登録されたアプリをリストボックスに表示
    def disp_app_info(self):
        # リストボックスをクリア
        self.frame_drag_drop.listbox.delete(0, tk.END)

        # テキストボックスにDBから読み込んだパスを表示する。DBが作成されていなければデフォルト値を表示する
        # DBからアプリ情報を読み込む
        self.cur.execute('SELECT name, path FROM apps ORDER BY id ASC')
        apps = self.cur.fetchall()
        if not apps:
            self.frame_drag_drop.listbox.insert(tk.END, "ここにファイルをドロップ")
        else:
            for app in apps:
                self.frame_drag_drop.listbox.insert(tk.END, f'{app[0]}\n')

        self.frame_drag_drop.listbox.see(tk.END)

    # アプリ終了時、DBを切断
    def __del__(self):
        self.conn.close()


def main():
    app = DragAndDrop()
    app.mainloop()


if __name__ == "__main__":
    main()
