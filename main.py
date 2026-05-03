"""Точка входа в приложение. Выполняет проверку зависимостей и запуск GUI."""
import sys
import tkinter as tk

try:
    import pandas as pd
    import openpyxl
    import xlrd
except ImportError as e:
    print(f"❌ Отсутствуют необходимые библиотеки. Установите их командой:")
    print("   pip install pandas openpyxl xlrd")
    sys.exit(1)

from gui.app import ArticleFinderGUI


def main() -> None:
    """Инициализирует главное окно и запускает цикл событий Tkinter."""
    root = tk.Tk()
    app = ArticleFinderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()