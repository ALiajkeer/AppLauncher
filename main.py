import tkinter as tk
from tkinter import simpledialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import sqlite3
import os
import subprocess
import pystray
from PIL import Image
import threading
import logging

APP_DEF_WIDTH = 400
APP_DEF_HEIGHT = 200
LIST_ITEMS = 10


# 登録アプリ表示関連
class AppList(tk.LabelFrame):
    def __init__(self, root):
        try:
            super().__init__(root, bd=2, relief="ridge", text="登録アプリ")
            self.grid(row=0, sticky="we")
            self.root = root

            # アプリ情報格納用
            self.app_dict = {}

            # DBに接続し、テーブルを作成。既にテーブルが存在するなら作成しない。
            self.conn = sqlite3.connect('app.db')
            self.cur = self.conn.cursor()
            self.cur.execute('''CREATE TABLE IF NOT EXISTS apps
                               (id INTEGER PRIMARY KEY,
                                name TEXT,
                                path TEXT,
                                icon_path TEXT)''')
            self.conn.commit()

            # リストボックス
            self.listbox = tk.Listbox(self, width=60, height=LIST_ITEMS, listvariable=tk.StringVar(), selectmode="single")
            self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            # スクロールバー
            self.scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.listbox.yview)
            self.scrollbar.pack(side=tk.LEFT, fill=tk.Y)
            # リストボックスとスクロールバーを連動させる
            self.listbox['yscrollcommand'] = self.scrollbar.set

            # DBに登録しているアプリ情報を表示
            self.disp_app_info()

            # ドラッグアンドドロップ
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind('<<Drop>>', self.func_drag_and_drop)

            # アプリをダブルクリック、もしくはエンターキーで起動
            self.listbox.bind('<Double-Button-1>', self.launch_app)
            self.listbox.bind("<Return>", self.launch_app)

            # 「DEL」キーでアプリを削除
            self.listbox.bind('<Delete>', self.delete_selected_item)

            # 右ダブルクリックで、アプリの名前を変更
            self.listbox.bind('<Double-Button-3>', self.on_listbox_doubleclick)
        except Exception as e:
            logging.exception("アプリ表示関数の初期化処理で異常発生: %s", e)

    # DBへ保存
    def save_app_info(self, name, path, icon_path):
        try:
            # データベースにアプリ情報を保存
            self.cur.execute("INSERT INTO apps(name, path, icon_path) VALUES (?, ?, ?)", (name, path, icon_path))
            self.conn.commit()
        except Exception as e:
            logging.exception("データベース保存処理で異常発生: %s", e)

    # アプリ情報をDBから削除
    def delete_selected_item(self, event):
        try:
            # 選択されているアイテムを取得
            selected_item = self.listbox.curselection()
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
        except Exception as e:
            logging.exception("登録アプリ削除処理で異常発生: %s", e)

    # ドラッグアンドドロップ時、アプリを登録する
    def func_drag_and_drop(self, event):
        try:
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
                messagebox.showerror("警告", "exeファイル以外のファイルは登録できません。")
        except Exception as e:
            logging.exception("アプリ登録処理で異常発生: %s", e)

    # DBに登録されたアプリをリストボックスに表示
    def disp_app_info(self):
        try:
            # リストボックスをクリア
            self.listbox.delete(0, tk.END)

            # テキストボックスにDBから読み込んだパスを表示する。DBが作成されていなければデフォルト値を表示する
            # DBからアプリ情報を読み込む
            self.cur.execute('SELECT id, name, path FROM apps ORDER BY id ASC')
            apps = self.cur.fetchall()
            # アプリ情報格納用辞書を空にする
            self.app_dict.clear()
            # データベースの登録が無い場合は「ここにファイルをドロップ」と表示する
            if not apps:
                self.listbox.insert(tk.END, "ここにファイルをドロップ")
            else:
                for app in apps:
                    # 辞書に登録しつつリストボックスに追加
                    self.app_dict.update({app[1]: app[2]})
                    self.listbox.insert(tk.END, f'{app[1]}\n')
            self.listbox.see(tk.END)
        except Exception as e:
            logging.exception("アプリ表示処理で異常発生: %s", e)

    # リストボックスのアプリをダブルクリックで起動
    def launch_app(self, event):
        try:
            # 選択されたアイテムのテキストを取得し、改行を削除
            selected_text = self.listbox.get(self.listbox.curselection()).rstrip()
            # アイテムのフルパスを取得
            full_path = self.app_dict.get(selected_text)
            # アプリのパスを辞書から取得  # アプリのディレクトリを取得
            app_dir = os.path.dirname(str(full_path))
            # アプリを起動  # 実行時の作業フォルダをアプリのディレクトリに設定
            if str(full_path).lower().endswith(".exe"):
                subprocess.run(str(full_path).encode('utf-8'), cwd=app_dir)
            # ショートカットの場合
            elif str(full_path).lower().endswith(".lnk"):
                subprocess.run(f"start /B {full_path}", shell=True)
            else:
                messagebox.showerror("警告", "有効なファイルではありません")
        except Exception as e:
            logging.exception("登録アプリ起動で異常発生: %s", e)

    # 右ダブルクリックで、アプリの名前を変更
    def on_listbox_doubleclick(self, event):
        try:
            # ダブルクリックされたアプリの名前を取得
            selection = self.listbox.curselection()
            if not selection:
                return
            name = self.listbox.get(selection[0]).strip()
            # 名前変更ダイアログを表示し、新しい名前を取得
            new_name = simpledialog.askstring("名前変更", f"{name} の新しい名前を入力してください", initialvalue=name)
            # キャンセルが押された場合は何もしない
            if new_name is None:
                return
            # 新しい名前をDBに反映させる
            self.cur.execute("UPDATE apps SET name = ? WHERE name = ?", (new_name, name))
            self.conn.commit()
            # リストボックスの表示を更新
            self.disp_app_info()
        except Exception as e:
            logging.exception("アプリの名前変更で異常発生: %s", e)

    # アプリ終了時、DBを切断
    def __del__(self):
        self.conn.close()


# タスクトレイ関連
class TaskTray:
    def __init__(self, root):
        try:
            self.root = root
            # アイコン画像を用意する
            # self.icon_image = Image.open("./Image/icon-48.ico")
            self.icon_image = Image.open('./icon-48.ico')
            self.icon = None
        except Exception as e:
            logging.exception("タスクトレイ関数の初期化処理で異常発生: %s", e)

    # タスクトレイ表示
    def start_icon_thread(self):
        try:
            # タスクトレイアイコンを作成する
            menu = pystray.Menu(pystray.MenuItem('Open', self.toggle_window, default=True))
            self.icon = pystray.Icon('app_name', self.icon_image, 'App Name', menu)
            # タスクトレイアイコンを表示する
            self.icon.run()
        except Exception as e:
            logging.exception("タスクトレイ表示スレッドで異常発生: %s", e)

    # タスクトレイアイコンをクリックするたびに、メインウィンドウの表示/非表示を切り替える。
    def toggle_window(self, *args):
        try:
            if self.root.state() == "normal":
                self.root.withdraw()
            else:
                self.root.deiconify()
        except Exception as e:
            logging.exception("メインウィンドウの表示/非表示処理で異常発生: %s", e)

    # アプリ終了時、タスクトレイアイコンを削除する
    def stop_icon_thread(self):
        try:
            if self.icon is not None:
                self.icon.stop()
        except Exception as e:
            logging.exception("タスクトレイ表示スレッド停止処理で異常発生: %s", e)


# ×ボタンでアプリ終了(タスクトレイアイコンを停止させる)
def exit_app(root, task_tray):
    try:
        # タスクトレイアイコンを停止
        task_tray.stop_icon_thread()
        # メインウィンドウを非表示にする
        root.withdraw()
        # メインウィンドウを破棄する
        root.destroy()
    except Exception as e:
        logging.exception("×ボタンでのアプリ終了処理で異常発生: %s", e)


def main():
    # インスタンス作成
    root = TkinterDnD.Tk()
    task_tray = TaskTray(root)

    # ログレベル設定
    logging.basicConfig(level=logging.INFO, filename='app.log', format='%(asctime)s %(levelname)s %(message)s',)

    # ウィンドウサイズ、タイトルの設定
    root.geometry(f'{APP_DEF_WIDTH}x{APP_DEF_HEIGHT}')
    root.minsize(APP_DEF_WIDTH, APP_DEF_HEIGHT)
    root.maxsize(APP_DEF_WIDTH + 100, APP_DEF_HEIGHT + 100)
    root.title(f'アプリランチャー')
    root.protocol("WM_DELETE_WINDOW", lambda: exit_app(root, task_tray))

    # 登録アプリをリストに表示
    app = AppList(root=root)

    # 別スレッドでアイコン表示を開始する
    icon_thread = threading.Thread(target=task_tray.start_icon_thread)
    icon_thread.start()

    app.mainloop()


# 直接呼び出し
main()
