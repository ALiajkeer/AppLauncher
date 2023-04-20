import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import sqlite3
import os
import subprocess
import pystray
from PIL import Image
import threading

APP_DEF_WIDTH = 300
APP_DEF_HEIGHT = 100


class AppList(TkinterDnD.Tk):
    def __init__(self, frame_drag_drop):
        # アプリ情報格納用
        self.app_dict = {}

        # メイン関数からフレームを受け継ぐ
        self.frame_drag_drop = frame_drag_drop

        # DBに接続し、テーブルを作成。既にテーブルが存在するなら作成しない。
        self.conn = sqlite3.connect('app.db')
        self.cur = self.conn.cursor()
        self.cur.execute('''CREATE TABLE IF NOT EXISTS apps
                           (id INTEGER PRIMARY KEY,
                            name TEXT,
                            path TEXT,
                            icon_path TEXT)''')
        self.conn.commit()

    # DBへ保存
    def save_app_info(self, name, path, icon_path):
        # データベースにアプリ情報を保存
        self.cur.execute("INSERT INTO apps(name, path, icon_path) VALUES (?, ?, ?)", (name, path, icon_path))
        self.conn.commit()

    # アプリ情報をDBから削除
    def delete_selected_item(self, event):
        # 選択されているアイテムを取得
        selected_item = self.frame_drag_drop.listbox.curselection()
        if not selected_item:
            return
        item_id = int(selected_item[0]) + 1
        # データベースからアプリ情報を削除
        self.cur.execute(f'DELETE FROM apps WHERE id={item_id}')
        self.conn.commit()

        # データベースのIDを振りなおす
        self.cur.execute("SELECT id FROM apps ORDER BY id")
        results = self.cur.fetchall()
        if results:
            # idを振り直す
            for i, r in enumerate(results):
                self.cur.execute(f"UPDATE apps SET id={i + 1} WHERE id={r[0]}")
                self.conn.commit()

        # アプリ情報を再表示
        self.disp_app_info()

    # ドラッグアンドドロップ時、ファイルのパスを取得する
    def func_drag_and_drop(self, event):
        # ドロップされたファイルからアプリ情報を取得
        file_path = event.data.strip()
        # 拡張子なしのファイル名を抽出
        name = os.path.splitext(os.path.basename(file_path))
        full_path = event.data.strip('{}\'')
        icon_path = ''  # アイコンファイルパスはとりあえず空にしておく

        # exeファイルのみDBへ登録
        if full_path.lower().endswith(".exe"):
            # アプリ情報をデータベースに保存
            self.save_app_info(name[0], full_path, icon_path)
            # リストボックスに表示
            self.disp_app_info()
        else:
            self.frame_drag_drop.messagebox.showwarning("警告", "exeファイル以外のファイルは登録できません。")

    # DBに登録されたアプリをリストボックスに表示
    def disp_app_info(self):
        # リストボックスをクリア
        self.frame_drag_drop.listbox.delete(0, tk.END)

        # テキストボックスにDBから読み込んだパスを表示する。DBが作成されていなければデフォルト値を表示する
        # DBからアプリ情報を読み込む
        self.cur.execute('SELECT id, name, path FROM apps ORDER BY id ASC')
        apps = self.cur.fetchall()
        # アプリ情報格納用辞書を空にする
        self.app_dict.clear()
        # データベースの登録が無い場合は「ここにファイルをドロップ」と表示する
        if not apps:
            self.frame_drag_drop.listbox.insert(tk.END, "ここにファイルをドロップ")
        else:
            for app in apps:
                # 辞書に登録しつつリストボックスに追加
                self.app_dict.update({app[1]: app[2]})
                self.frame_drag_drop.listbox.insert(tk.END, f'{app[1]}\n')
        self.frame_drag_drop.listbox.see(tk.END)

    # リストボックスのアプリをダブルクリックで起動
    def launch_app(self, event):
        # 選択されたアイテムのテキストを取得し、改行を削除
        selected_text = self.frame_drag_drop.listbox.get(self.frame_drag_drop.listbox.curselection()).rstrip()
        # アプリのパスを辞書から取得  # アプリのディレクトリを取得
        app_dir = os.path.dirname(self.app_dict.get(selected_text))
        # アプリを起動  # 実行時の作業フォルダをアプリのディレクトリに設定
        subprocess.run([self.app_dict.get(selected_text)], cwd=app_dir)

    # アプリ終了時、DBを切断
    def __del__(self):
        self.conn.close()


# タスクトレイ関連
class TaskTray:
    def __init__(self, root):
        self.root = root
        # アイコン
        # self.icon = ('icon.ico', 'icon_16x16.ico', 'icon_32x32.ico')

    # タスクトレイ表示
    def start_icon_thread(self):
        # アイコン画像を用意する
        icon_image = Image.open("./Image/icon-16.png")
        # タスクトレイアイコンを作成する
        menu = pystray.Menu(pystray.MenuItem('Open', self.toggle_window, default=True))
        icon = pystray.Icon('app_name', icon_image, 'App Name', menu)
        # タスクトレイアイコンを表示する
        icon.run()

    # タスクトレイアイコンをクリックするたびに、メインウィンドウの表示/非表示を切り替える。
    def toggle_window(self, *args):
        if self.root.state() == "normal":
            self.root.withdraw()
        else:
            self.root.deiconify()


def main():
    # インスタンス作成
    root = TkinterDnD.Tk()

    # ウィンドウサイズ、タイトルの設定
    root.geometry(f'{APP_DEF_WIDTH}x{APP_DEF_HEIGHT}')
    root.minsize(APP_DEF_WIDTH, APP_DEF_HEIGHT)
    root.maxsize(APP_DEF_WIDTH + 100, APP_DEF_HEIGHT + 100)
    root.title(f'アプリランチャー')

    # フレーム
    # ラベルフレームの作成（ラベルフレームのtextをmenuに設定）
    frame_drag_drop = tk.LabelFrame(root, bd=2, relief="ridge", text="menu")
    frame_drag_drop.grid(row=0, sticky="we")

    # リストボックス
    frame_drag_drop.listbox = tk.Listbox(frame_drag_drop, height=5, listvariable=tk.StringVar(), selectmode="single")
    frame_drag_drop.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # スクロールバー
    frame_drag_drop.scrollbar = tk.Scrollbar(frame_drag_drop, orient=tk.VERTICAL, command=frame_drag_drop.listbox.yview)
    frame_drag_drop.scrollbar.pack(side=tk.LEFT, fill=tk.Y)
    # リストボックスとスクロールバーを連動させる
    frame_drag_drop.listbox['yscrollcommand'] = frame_drag_drop.scrollbar.set

    # DBに登録しているアプリ情報を表示
    drag_and_drop = AppList(frame_drag_drop)
    drag_and_drop.disp_app_info()

    # ドラッグアンドドロップ
    frame_drag_drop.listbox.drop_target_register(DND_FILES)
    frame_drag_drop.listbox.dnd_bind('<<Drop>>', drag_and_drop.func_drag_and_drop)

    # アプリをダブルクリックで起動
    frame_drag_drop.listbox.bind('<Double-Button-1>', drag_and_drop.launch_app)

    # 「DEL」キーでアプリを削除
    frame_drag_drop.listbox.bind('<Delete>', drag_and_drop.delete_selected_item)

    # 配置
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # 別スレッドでアイコン表示を開始する
    task_tray = TaskTray(root)
    icon_thread = threading.Thread(target=task_tray.start_icon_thread)
    icon_thread.start()

    root.mainloop()


if __name__ == "__main__":
    main()
