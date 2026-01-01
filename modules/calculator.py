import tkinter as tk

from core.module_loader import BaseModule


class CalculatorModule(BaseModule):
    """
    Демонстрационный модуль - калькулятор, умеющий только умножать.
    """
    name = "calculator"
    button_label = "Модуль для быстрых вычислений"
    button_text = "Запустить калькулятор"
    module_label = (
        "Описание функционала модуля - калькулятор, умеющий только умножать"
    )

    @classmethod
    def initialize_frame(cls, parent_frame: tk.Frame) -> None:
        input_var1 = tk.StringVar(value="0")
        input_var2 = tk.StringVar(value="0")
        result_var = tk.StringVar(value="Ожидаю ввода...")

        tk.Label(parent_frame, text="Первое число:").pack()
        tk.Entry(parent_frame, textvariable=input_var1).pack(pady=5)

        tk.Label(parent_frame, text="Второе число:").pack()
        tk.Entry(parent_frame, textvariable=input_var2).pack(pady=5)

        tk.Button(
            parent_frame,
            text="Умножить",
            command=lambda: result_var.set(
                f"Произведение: {float(input_var1.get()) * float(
                    input_var2.get())}"
                )
            ).pack(pady=10)

        tk.Label(parent_frame, textvariable=result_var).pack()
