# app/config/app_config.py

from pathlib import Path


# Определяем корневую директорию проекта.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
MODULES_DIR = PROJECT_ROOT / "modules"
CONFIG_PATHS_NAME = PROJECT_ROOT / "config" / "config_paths_name.toml"

# Название приложения, логотив в строке заголовка.
APP_TITLE = "The F.I.R. 2.0a (File Interaction Runner)"
APP_LOGO = PROJECT_ROOT / "images" / "logo_small.png"

# Определяем геометрию приложения.
GEOMETRY_MAIN_WINDOW = "700x300"  # Начальный/минимальный размер окна.
MAIN_WINDOW_MAXSIZE = "1500x900"  # Максимальный размер окна.
WRAPLENGHT_BUTTON = 200  # Максимальная ширина лейбла, над кнопкой модуля.

# Определяем, сколько подключаемых модулей может быть открыто одновременно.
MAX_CONCURRENT_MODULES = 3
