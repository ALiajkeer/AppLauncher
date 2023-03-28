import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD


class MyApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # ウィンドウサイズ
        width = 300
        height = 100
        self.geometry(f'{width}x{height}')
        self.minsize(width, height)
        self.maxsize(width, height)
        self.title(f'DnD')

        # フレーム
        self.frame_drag_drop = DragAndDrop(self)

        # 配置
        self.frame_drag_drop.grid(column=0, row=0, padx=5, pady=5, sticky=(tk.E, tk.W, tk.S, tk.N))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)


class DragAndDrop(tk.LabelFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.textbox = tk.Text(self)
        self.textbox.insert(0.0, "Drag and Drop")
        self.textbox.configure(state='disabled')

        # ドラッグアンドドロップ
        self.textbox.drop_target_register(DND_FILES)
        self.textbox.dnd_bind('<<Drop>>', self.funcDragAndDrop)

        # スクロールバー設定
        self.scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.textbox.yview)
        self.textbox['yscrollcommand'] = self.scrollbar.set

        # 配置
        self.textbox.grid(column=0, row=0, sticky=(tk.E, tk.W, tk.S, tk.N))
        self.scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)


    def funcDragAndDrop(self, e):
        # ここを編集してください
        message = '\n' + e.data

        self.textbox.configure(state='normal')
        self.textbox.insert(tk.END, message)
        self.textbox.configure(state='disabled')

        self.textbox.see(tk.END)


if __name__ == "__main__":
    app = MyApp()
    app.mainloop()
