import tkinter as tk
import random
import playsound
import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror

class SpellingApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry('300x500')

        # 组件初始化
        self.word_label = tk.Label(root)
        self.word_label.pack()

        self.pronunciation_label = tk.Label(root)
        self.pronunciation_label.pack()

        self.entry = tk.Entry(root)
        self.entry.pack()
        self.entry.bind("<Return>", lambda e: self.check_spelling())

        self.check_button = tk.Button(root, text="Check Spelling", command=self.check_spelling)
        self.check_button.pack()

        self.play_audio_button = tk.Button(root, text="Play Audio", command=self.play_audio)
        self.play_audio_button.pack()

        self.mode_var = tk.StringVar(root, value="sequential")
        self.mode_option = tk.OptionMenu(root, self.mode_var, "sequential", "random", "reverse")
        self.mode_option.pack()

        self.prev_button = tk.Button(root, text="Previous Word", command=self.previous_word)
        self.prev_button.pack()

        self.next_button = tk.Button(root, text="Next Word", command=self.next_word)
        self.next_button.pack()

        self.indicator_label = tk.Label(root, text="")
        self.indicator_label.pack()

        self.history_text = tk.Text(root)
        self.history_text.pack()
        self.history_text.tag_config('incorrect', foreground='red')

        self.save_button = tk.Button(root, text="Save History", command=self.save_history)
        self.save_button.pack()

        self.words = []
        self.current_word = []
        self.word_index = 0

    def load_words(self):
        filename = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not filename:
            return
        self.words = pd.read_excel(filename).values.tolist()
        self.words = [(str(w[0]), str(w[1]), str(w[2])) for w in self.words]
        if self.mode_var.get() == "random":
            random.shuffle(self.words)
        elif self.mode_var.get() == "reverse":
            self.words = list(reversed(self.words))
        self.word_index = 0
        self.display_current_word()

    def display_current_word(self):
        self.current_word = self.words[self.word_index]
        self.word_label.config(text=self.current_word[2])
        self.pronunciation_label.config(text=self.current_word[1])
        self.entry.delete(0, tk.END)

    def check_spelling(self):
        if self.entry.get() == self.current_word[0]:
            self.indicator_label.config(text="Correct", bg="green")
            self.history_text.insert(tk.END, self.current_word[0] + ' - ' + self.current_word[2] + '\n')
            self.next_word()
        else:
            self.indicator_label.config(text="Incorrect", bg="red")
            self.history_text.insert(tk.END, self.current_word[0] + ' - ' + self.current_word[2] + ' (incorrect)\n', 'incorrect')

    def previous_word(self):
        self.word_index = max(0, self.word_index - 1)
        self.display_current_word()

    def next_word(self):
        self.word_index = min(len(self.words) - 1, self.word_index + 1)
        self.display_current_word()

    def play_audio(self):
        audio_file = "Sounds/Speech_US/" + self.current_word[0] + ".mp3"
        playsound.playsound(audio_file, True)

    def save_history(self):
        with open("history.md", "w") as f:
            f.write(self.history_text.get("1.0", tk.END))

