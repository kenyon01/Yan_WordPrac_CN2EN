import tkinter as tk
from gui import SpellingApp

def main():
    root = tk.Tk()
    app = SpellingApp(root)
    app.load_words()  # 从文件加载单词
    root.mainloop()

if __name__ == "__main__":
    main()
