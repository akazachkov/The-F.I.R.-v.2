"""
Песочница, в которой реализован интерфейс с двумя фреймами, которые отличаются
по параметрам.
"""

import tkinter as tk
from tkinter import ttk


def create_simple_dynamic_frame(master, title, items, width=None):
    """
    Создает фрейм с динамическими шириной и высотой.

    Ширина зависит от ширины основного окна.
    Высота зависит от количества содержимого фрейма.
    """
    frame = ttk.Frame(master, relief="solid", padding=5)
    frame.pack(fill='x', pady=5)

    # Рассчитываем wraplength, если ширина задана
    wrap_len = (width - 10) if width else 0

    ttk.Label(
        frame,
        text=title,
        font=("Arial", 12, "bold"),
        wraplength=wrap_len,  # Перенос длинного текста на новую строку
        justify='left'
        ).pack(pady=5, anchor='w')

    for item in items:
        ttk.Checkbutton(frame, text=item).pack(anchor='w', padx=5)

    return frame


def create_fixed_width_frame(master, title, items, width_frame):
    """
    Создает фрейм с фиксированной шириной и динамической высотой.

    Ширина задаётся в параметрах фрейма.
    Высота зависит от количества содержимого фрейма.
    """
    # ШАГ 1: Внешний контейнер для центрирования
    # Он растянется на всю ширину, но внутри себя будет центрировать фрейм.
    container = tk.Frame(master)
    container.pack(fill='x', pady=5)

    # ШАГ 2: Сам фрейм с фиксированной шириной
    fixed_frame = tk.Frame(
        container,
        width=width_frame,
        bd=1,
        relief="solid",
        bg="white"
        )

    # ШАГ 3: Центрируем с помощью pack
    # expand=False - не расширяться
    # anchor='nw' - положение в верхнем левом углу, чтобы не было лишнего
    # вертикального пространства
    fixed_frame.pack(expand=False, anchor='nw')

    # ШАГ 4: Отключаем подстройку под размер родителя (container)
    fixed_frame.pack_propagate(False)

    # ШАГ 5: Наполняем фрейм и замеряем нужную высоту
    content = tk.Frame(fixed_frame, bg="white")
    content.pack(fill='both', expand=True, padx=5, pady=10)
    content.columnconfigure(0, weight=1)

    # Заголовок
    lbl_title = tk.Label(
        content,
        text=title,
        font=("Arial", 11, "bold"),
        wraplength=width_frame,
        justify='left',
        bg="white"
        )
    lbl_title.grid(row=0, column=0, sticky='ew', pady=2)

    # Чекбоксы
    for i, item in enumerate(items, 1):
        chk = tk.Checkbutton(
            content,
            text=item,
            wraplength=width_frame - 15,
            anchor='w',
            justify='left',
            bg="white"
            )
        chk.grid(row=i, column=0, sticky='w')

    # ШАГ 6: Ручная установка высоты
    fixed_frame.update_idletasks()
    required_height = content.winfo_reqheight() + 20
    fixed_frame.config(height=required_height)

    return container


# --- ГЛАВНЫЙ БЛОК ---
if __name__ == '__main__':
    root = tk.Tk()
    root.title("Основное окно")
    root.geometry("500x650")

    super_long_title_1 = (
        "Динамический фрейм, подстраивающийся по высоте (в зависимости от "
        "содержимого) и ширине (в зависимости от ширины основного окна)"
        )
    items_list_1 = ['Item 1', 'Item 2', 'Item 3', 'Item 4', 'Item 5',
                    'Item 6', 'Item 7', 'Item 8', 'Item 9', 'Item 10']
    super_long_title_2 = (
        "Второй фрейм (фиксированная ширина) с длинным заголовком, который "
        "будет автоматически перенесен на новую строку в текущем фрейме"
        )
    items_list_2 = [
        'Option A',
        ('Option B - Длинный текст в опции, который тоже будет автоматически'
            ' перенесен на новую строку в текущем фрейме'),
        'Option C',
        'Option D',
        'Option E',
        'Option F'
        ]

    # --- Фрейм 1 ---
    create_simple_dynamic_frame(
        root,
        super_long_title_1,
        items_list_1,
        width=350  # Ширина фрейма титула, для переноса длинного текста
        )

    # --- Фрейм 2 (с длинным заголовком) ---
    create_fixed_width_frame(
        root,
        super_long_title_2,
        items_list_2,
        width_frame=350  # Ширина всего фрейма
    )

    root.mainloop()
