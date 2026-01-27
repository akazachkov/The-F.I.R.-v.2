# app/core/elements/working_with_folders.py

import subprocess
from pathlib import Path


def ensure_directory_exists(base_path: Path, subfolder: str) -> Path:
    """
    Создает директорию, если она не существует.

    Args:
        base_path - Путь к базовой директории.
        subfolder - Имя новой поддиректории.
    """
    directory = base_path / subfolder
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def open_file_and_folder(path: str | Path) -> bool:
    """
    Открывает указанную директорию или файл в Проводнике Windows.

    Args:
        path - Строка или Path-объект, указывающий на директорию.

    Returns:
        bool - True, если команда была успешно запущена, иначе False.
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
