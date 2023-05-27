import pandas as pd
import random
from playsound import playsound
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QLineEdit, QComboBox, QTextEdit, QFileDialog
from PyQt5.QtCore import Qt

class SpellingApp(QWidget):
    def __init__(self):
        super().__init__()
        self.words = []
        self.current_word = None
        self.current_index = -1

        self.init_ui()

    def init_ui(self):
        self.resize(300, 500)

        self.layout = QVBoxLayout()

        self.word_label = QLabel()
        self.word_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.word_label)

        self.pronunciation_label = QLabel()
        self.pronunciation_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.pronunciation_label)

        self.entry = QLineEdit()
        self.entry.returnPressed.connect(self.check_spelling)
        self.layout.addWidget(self.entry)

        self.test_mode = QComboBox()
        self.test_mode.addItem("顺序测试")
        self.test_mode.addItem("随机模式")
        self.test_mode.addItem("逆序测试")
        self.layout.addWidget(self.test_mode)

        self.check_button = QPushButton("检查拼写")
        self.check_button.clicked.connect(self.check_spelling)
        self.layout.addWidget(self.check_button)

        self.play_sound_button = QPushButton("播放音频")
        self.play_sound_button.clicked.connect(self.play_audio)
        self.layout.addWidget(self.play_sound_button)

        self.next_word_button = QPushButton("下一个单词")
        self.next_word_button.clicked.connect(self.next_word)
        self.layout.addWidget(self.next_word_button)

        self.prev_word_button = QPushButton("上一个单词")
        self.prev_word_button.clicked.connect(self.prev_word)
        self.layout.addWidget(self.prev_word_button)

        self.history = QTextEdit()
        self.history.setReadOnly(True)
        self.layout.addWidget(self.history)

        self.setLayout(self.layout)

    def load_words(self, filename):
        df = pd.read_excel(filename)
        self.words = df.values.tolist()
        self.next_word()

    def next_word(self):
        self.entry.clear()
        self.current_index += 1
        if self.test_mode.currentText() == "随机模式":
            self.current_index = random.randint(0, len(self.words) - 1)
        elif self.test_mode.currentText() == "逆序测试":
            self.current_index = len(self.words) - 1 - self.current_index
        self.current_word = self.words[self.current_index]
        self.word_label.setText(self.current_word[2])
        self.pronunciation_label.setText(self.current_word[1])

    def prev_word(self):
        if self.current_index > 0:
            self.current_index -= 1
        self.current_word = self.words[self.current_index]
        self.word_label.setText(self.current_word[2])
        self.pronunciation_label.setText(self.current_word[1])
        
    def check_spelling(self):
        if self.entry.text() == self.current_word[0]:
            self.history.append(f"{self.current_word[0]}: {self.current_word[2]}")
            self.next_word()
        else:
            self.history.append(f"<font color='red'>{self.current_word[0]}: {self.current_word[2]}</font>")

    def play_audio(self):
        audio_file = "Sounds/Speech_US/" + self.current_word[0] + ".mp3"
        playsound(audio_file)
