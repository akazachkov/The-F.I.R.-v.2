# app/core/elements/open_file_or_folder.py

import subprocess
from pathlib import Path


# Функция для открытия файла и директории (универсальная)
def open_file_and_folder(path: str | Path) -> bool:
    """
    Открывает указанную директорию или файл в Проводнике Windows.

    Args:
        path: Строка или Path-объект, указывающий на директорию.

    Returns:
        bool: True, если команда была успешно запущена, иначе False.
    """
    try:
        # Нормализуем путь, преобразуя его в Windows-совместимый вид
        target_path = Path(path).resolve()

        # Проверяем, существует ли путь
        if not target_path.exists():
            print(f"Ошибка: путь '{target_path}' не найден.")
            return False

        # Самая простая и надежная команда
        subprocess.run(["explorer", str(target_path)])
        return True

    except subprocess.CalledProcessError as e:
        print(f"Не удалось запустить 'explorer'. Ошибка: {e}")
        return False
    except Exception as e:
        print(f"Ошибка: {e}")
        return False
