# app/app.py

import sys
import traceback

from core.app_controller import AppController
from gui.main_window import MainWindow
from config.app_config import CONFIG_PATHS_NAME


def main():
    try:
        # Передаём путь к конфигурации при создании контроллера
        controller = AppController(config_path=CONFIG_PATHS_NAME)
        app = MainWindow(controller)

        # Запускаем GUI цикл
        app.mainloop()
    except Exception as e:
        print("Ошибка при запуске приложения:", e)
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    """Запускает приложение."""
    main()
