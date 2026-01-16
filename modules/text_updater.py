import tkinter as tk
import threading

from core.module_loader import BaseModule


class TextUpdaterModule(BaseModule):
    """
    Демонстрационный модуль - безопасная работа с потоками и GUI.
    """
    name = "text_updater"
    button_label = "Асинхронное обновление текстового поля"
    button_text = "Запустить текст-апдейтер"
    module_label = (
        "Описание функционала модуля - безопасная работа с потоками и GUI"
    )

    @classmethod
    def initialize_frame(cls, parent_frame: tk.Frame) -> None:
        info_label = tk.Label(parent_frame, text="Модуль TextUpdater активен.")
        info_label.pack(pady=10)

        text_widget = tk.Text(parent_frame, height=5, width=50)
        text_widget.pack(pady=5)

        def long_running_task():
            # Эта задача выполняется в фоновом потоке
            import time
            for i in range(5):
                time.sleep(1)
                # Для обновления GUI из потока мы используем gui_runner
                # gui_run должен быть доступен через cls
                try:
                    cls.gui_run(
                        lambda: text_widget.insert(
                            tk.END, f"Сообщение из потока {i+1}\n"
                        )
                    )
                except AttributeError:
                    # Запасной вариант, если gui_run не настроен
                    # Просто печатаем в консоль
                    print(f"Сообщение из потока {i+1}")

        start_button = tk.Button(
            parent_frame,
            text="Запустить фоновую задачу",
            command=lambda: threading.Thread(
                target=long_running_task, daemon=True).start()
        )
        start_button.pack(pady=10)
