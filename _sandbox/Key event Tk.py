import tkinter as tk
from tkinter import scrolledtext, ttk


class KeyEventApp:
    """
    Код для работы с сочетаниями клавиш на английской и русской раскладках,
    при использовании Tkinter.
    """
    def __init__(self, root):
        self.root = root
        self.root.geometry("600x400")

        # Словарь для преобразования кодов клавиш в символы
        self.virtual_keys = {
            65: 'A', 67: 'C', 83: 'S', 86: 'V', 88: 'X', 89: 'Y', 90: 'Z'}

        self.setup_ui()

    def setup_ui(self):
        # Создаем основную рамку
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid()

        # Метка и поле для ввода текста
        input_label = ttk.Label(main_frame)
        input_label.grid()

        self.text_input = scrolledtext.ScrolledText(
            main_frame, wrap=tk.WORD, height=10)
        self.text_input.grid()

        # Привязываем обработчики событий
        self.text_input.bind('<Control-KeyPress>', self.on_control_key)

    def on_control_key(self, event):
        """Обработчик нажатия клавиш с Ctrl"""
        keycode = event.keycode

        # Преобразуем код клавиши в букву для сочетания
        key_char = self.virtual_keys.get(keycode, '').upper()

        if key_char:
            # Выполняем действие в зависимости от сочетания
            if key_char == 'C':
                self.text_input.event_generate('<<Copy>>')
            elif key_char == 'V':
                self.text_input.event_generate('<<Paste>>')
            elif key_char == 'X':
                self.text_input.event_generate('<<Cut>>')
            elif key_char == 'A':
                self.text_input.event_generate('<<SelectAll>>')
            elif key_char == 'Z':
                self.text_input.event_generate('<<Undo>>')
            elif key_char == 'Y':
                self.text_input.event_generate('<<Redo>>')

            # Предотвращаем дальнейшую обработку
            return "break"

        return None


if __name__ == "__main__":
    root = tk.Tk()

    app = KeyEventApp(root)

    root.mainloop()
