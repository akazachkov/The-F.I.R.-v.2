import tkinter as tk
from tkinter import scrolledtext


"""
Оконное приложение, которое поможет определить скан-код клавиш или сочетания
клавиш, в конкретной системе.
"""


def update_key_info(event):
    # Очищаем предыдущую информацию
    output_field.config(state=tk.NORMAL)
    output_field.delete(1.0, tk.END)

    # Модификаторы
    modifiers = []
    if event.state & 0x1:  # Shift
        modifiers.append("Shift")
    if event.state & 0x4:  # Ctrl
        modifiers.append("Ctrl")
    if event.state & 0x8:  # Alt
        modifiers.append("Alt")
    if event.state & 0x20:  # Caps Lock
        modifiers.append("CapsLock")
    if event.state & 0x40:  # Num Lock
        modifiers.append("NumLock")

    # Основная клавиша
    key = event.keysym
    keycode = hex(event.keycode)

    # Формируем вывод
    output_field.insert(tk.END, f"Сочетание: {'+'.join(modifiers)}+{key}\n")
    output_field.insert(tk.END, f"Скан-код: {keycode}\n")
    output_field.insert(tk.END, f"State: {hex(event.state)}\n")
    output_field.insert(tk.END, "---" + "\n")

    output_field.config(state=tk.DISABLED)


# Создаем основное окно
root = tk.Tk()
root.title("Key Event Debugger")

# Создаем поле для вывода
output_field = scrolledtext.ScrolledText(
    root, height=15, width=50, state=tk.DISABLED)
output_field.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Привязываем обработчик ко всем нажатиям клавиш
root.bind("<KeyPress>", update_key_info)

# Добавляем инструкцию
instruction = tk.Label(root, text="Нажмите любую клавишу или сочетание клавиш")
instruction.pack(pady=5)

# Кнопка для очистки поля
clear_button = tk.Button(
    root,
    text="Очистить",
    command=lambda: (
        output_field.config(state=tk.NORMAL),
        output_field.delete(1.0, tk.END),
        output_field.config(state=tk.DISABLED)
    )
)
clear_button.pack(pady=5)

root.mainloop()
