"""
Приложение для поиска совпадающих артикулов в двух Excel файлах.
Архитектура построена по принципам SOLID:
- SRP: каждый класс отвечает за одну задачу (загрузка, парсинг, сравнение, GUI)
- OCP: расширение функционала (новые форматы, другие правила парсинга) не требует изменения существующих классов
- DIP: GUI зависит от абстракций (интерфейсов компонентов), а не от их конкретных реализаций
"""

import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Set, List

import pandas as pd


class ArticleExtractor:
    """
    Класс для извлечения артикулов из текстовых строк с помощью регулярного выражения.

    Parameters
    ----------
    pattern : str, optional
        Регулярное выражение для поиска артикулов. По умолчанию ищет последовательности
        букв (кириллица/латиница), цифр, дефисов и подчеркиваний длиной от 3 символов.
    """

    def __init__(self, pattern: str = r'[A-Za-zА-Яа-я0-9\-_]{3,}') -> None:
        self._pattern = re.compile(pattern)

    def extract(self, text: Optional[str]) -> Set[str]:
        """
        Извлекает уникальные артикулы из переданного текста.

        Parameters
        ----------
        text : str or None
            Текст для анализа. Может быть None, NaN или пустой строкой.

        Returns
        -------
        set of str
            Множество найденных уникальных артикулов.
        """
        if not isinstance(text, str) or not text.strip():
            return set()
        return set(self._pattern.findall(text))


class ExcelDataLoader:
    """
    Класс для загрузки и чтения данных из Excel файлов (.xls, .xlsx).

    Parameters
    ----------
    file_path : str
        Путь к Excel файлу.
    """

    def __init__(self, file_path: str) -> None:
        self._file_path = file_path
        self._dataframe: Optional[pd.DataFrame] = None

    def load(self) -> pd.DataFrame:
        """
        Загружает Excel файл в DataFrame с автоматическим выбором движка.

        Returns
        -------
        pd.DataFrame
            Загруженная таблица данных.

        Raises
        ------
        FileNotFoundError
            Если файл не существует.
        ValueError
            Если расширение файла не поддерживается.
        """
        if not os.path.exists(self._file_path):
            raise FileNotFoundError(f"Файл не найден: {self._file_path}")

        ext = os.path.splitext(self._file_path)[1].lower()
        if ext == ".xlsx":
            engine = "openpyxl"
        elif ext == ".xls":
            engine = "xlrd"
        else:
            raise ValueError("Поддерживаются только форматы .xls и .xlsx")

        self._dataframe = pd.read_excel(self._file_path, engine=engine)
        return self._dataframe

    def get_columns(self) -> List[str]:
        """
        Возвращает список названий колонок загруженного файла.

        Returns
        -------
        list of str
            Названия колонок. Если файл не загружен, возвращает пустой список.
        """
        if self._dataframe is None:
            return []
        return self._dataframe.columns.astype(str).tolist()

    def get_column_data(self, column_name: str) -> pd.Series:
        """
        Возвращает данные указанной колонки.

        Parameters
        ----------
        column_name : str
            Название колонки.

        Returns
        -------
        pd.Series
            Данные колонки.

        Raises
        ------
        ValueError
            Если данные не были загружены или колонка не найдена.
        """
        if self._dataframe is None:
            raise ValueError("Данные не загружены. Сначала вызовите метод load().")
        if column_name not in self._dataframe.columns:
            raise KeyError(f"Колонка '{column_name}' не найдена в файле.")
        return self._dataframe[column_name]


class ArticleMatcher:
    """Класс для поиска пересечений между двумя множествами артикулов."""

    def find_common(self, articles1: Set[str], articles2: Set[str]) -> Set[str]:
        """
        Находит артикулы, присутствующие в обоих множествах.

        Parameters
        ----------
        articles1 : set of str
            Первое множество артикулов.
        articles2 : set of str
            Второе множество артикулов.

        Returns
        -------
        set of str
            Множество общих артикулов.
        """
        return articles1.intersection(articles2)


class ArticleFinderGUI:
    """
    Графический интерфейс приложения для поиска совпадающих артикулов.

    Parameters
    ----------
    root : tk.Tk
        Корневое окно приложения Tkinter.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Поиск совпадений артикулов в Excel")
        self.root.geometry("750x650")
        self.root.resizable(False, False)

        # Внедрение зависимостей (Dependency Injection)
        self._extractor = ArticleExtractor()
        self._matcher = ArticleMatcher()
        self._loader1: Optional[ExcelDataLoader] = None
        self._loader2: Optional[ExcelDataLoader] = None

        self._setup_ui()

    def _setup_ui(self) -> None:
        """Инициализирует и размещает все элементы интерфейса."""
        style = ttk.Style()
        style.theme_use("clam")

        # --- Файл 1 ---
        frm1 = ttk.LabelFrame(self.root, text="Первый файл")
        frm1.pack(fill="x", padx=10, pady=5)

        self.path1_var = tk.StringVar()
        ttk.Entry(frm1, textvariable=self.path1_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(5, 0), pady=5
        )
        ttk.Button(frm1, text="Выбрать", command=lambda: self._select_file(self.path1_var)).pack(
            side="left", padx=5, pady=5
        )
        ttk.Button(frm1, text="Загрузить", command=self._load_file1).pack(
            side="left", padx=5, pady=5
        )

        self.col1_var = tk.StringVar()
        self.col1_cb = ttk.Combobox(frm1, textvariable=self.col1_var, state="readonly")
        self.col1_cb.pack(fill="x", padx=5, pady=(0, 5))

        # --- Файл 2 ---
        frm2 = ttk.LabelFrame(self.root, text="Второй файл")
        frm2.pack(fill="x", padx=10, pady=5)

        self.path2_var = tk.StringVar()
        ttk.Entry(frm2, textvariable=self.path2_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(5, 0), pady=5
        )
        ttk.Button(frm2, text="Выбрать", command=lambda: self._select_file(self.path2_var)).pack(
            side="left", padx=5, pady=5
        )
        ttk.Button(frm2, text="Загрузить", command=self._load_file2).pack(
            side="left", padx=5, pady=5
        )

        self.col2_var = tk.StringVar()
        self.col2_cb = ttk.Combobox(frm2, textvariable=self.col2_var, state="readonly")
        self.col2_cb.pack(fill="x", padx=5, pady=(0, 5))

        # --- Кнопка поиска ---
        ttk.Button(self.root, text="Найти совпадения", command=self._find_matches).pack(pady=10)

        # --- Результаты ---
        frm_res = ttk.LabelFrame(self.root, text="Найденные совпадения")
        frm_res.pack(fill="both", expand=True, padx=10, pady=5)

        self.res_text = tk.Text(frm_res, wrap="word", state="disabled", font=("Consolas", 10))
        scroll = ttk.Scrollbar(frm_res, command=self.res_text.yview)
        self.res_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.res_text.pack(side="left", fill="both", expand=True, padx=5, pady=5)

    def _select_file(self, string_var: tk.StringVar) -> None:
        """
        Открывает диалог выбора файла и сохраняет путь в переменную.

        Parameters
        ----------
        string_var : tk.StringVar
            Переменная для хранения пути к файлу.
        """
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if filepath:
            string_var.set(filepath)

    def _load_file1(self) -> None:
        """Загружает первый Excel файл и обновляет список колонок."""
        self._load_file(self.path1_var.get(), 1)

    def _load_file2(self) -> None:
        """Загружает второй Excel файл и обновляет список колонок."""
        self._load_file(self.path2_var.get(), 2)

    def _load_file(self, path: str, file_num: int) -> None:
        """
        Общая логика загрузки Excel файла.

        Parameters
        ----------
        path : str
            Путь к выбранному файлу.
        file_num : int
            Номер файла (1 или 2) для определения целевых переменных интерфейса.
        """
        if not path:
            messagebox.showwarning("Внимание", "Сначала выберите файл через диалог.")
            return

        try:
            loader = ExcelDataLoader(path)
            loader.load()
            cols = loader.get_columns()

            if file_num == 1:
                self._loader1 = loader
                self.col1_cb["values"] = cols
                self.col1_cb.current(0) if cols else self.col1_cb.set("")
            else:
                self._loader2 = loader
                self.col2_cb["values"] = cols
                self.col2_cb.current(0) if cols else self.col2_cb.set("")

            messagebox.showinfo("Успех", f"Файл {file_num} загружен. Доступно колонок: {len(cols)}")
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    def _find_matches(self) -> None:
        """Извлекает артикулы из выбранных колонок и выводит их пересечение."""
        if not self._loader1 or not self._loader2:
            messagebox.showwarning("Внимание", "Загрузите оба файла перед поиском.")
            return

        col1 = self.col1_var.get()
        col2 = self.col2_var.get()
        if not col1 or not col2:
            messagebox.showwarning("Внимание", "Выберите колонки для анализа в обоих файлах.")
            return

        try:
            self._update_status("Анализ первого файла...")
            articles1 = set()
            for val in self._loader1.get_column_data(col1).dropna():
                articles1.update(self._extractor.extract(str(val)))

            self._update_status("Анализ второго файла...")
            articles2 = set()
            for val in self._loader2.get_column_data(col2).dropna():
                articles2.update(self._extractor.extract(str(val)))

            self._update_status("Поиск совпадений...")
            common = self._matcher.find_common(articles1, articles2)

            self._display_results(common)
            messagebox.showinfo("Готово", f"Найдено совпадений: {len(common)}")
        except Exception as e:
            messagebox.showerror("Ошибка поиска", str(e))
        finally:
            self._update_status("Ожидание действий...")

    def _update_status(self, text: str) -> None:
        """
        Обновляет временное сообщение в поле результатов.

        Parameters
        ----------
        text : str
            Текст статуса.
        """
        self.res_text.config(state="normal")
        self.res_text.delete("1.0", tk.END)
        self.res_text.insert("1.0", text)
        self.res_text.config(state="disabled")
        self.root.update()

    def _display_results(self, matches: Set[str]) -> None:
        """
        Форматирует и выводит найденные совпадения в текстовое поле.

        Parameters
        ----------
        matches : set of str
            Множество найденных артикулов.
        """
        self.res_text.config(state="normal")
        self.res_text.delete("1.0", tk.END)

        if not matches:
            self.res_text.insert("1.0", "Совпадений не найдено.")
        else:
            self.res_text.insert("1.0", f"Найдено совпадений: {len(matches)}\n" + "-" * 30 + "\n")
            for art in sorted(matches):
                self.res_text.insert("end", f"• {art}\n")

        self.res_text.config(state="disabled")


if __name__ == "__main__":
    # Проверка наличия необходимых библиотек
    try:
        import pandas as pd
        import openpyxl  # noqa: F401
        import xlrd      # noqa: F401
    except ImportError as e:
        print(f"Отсутствуют необходимые библиотеки. Установите их командой:\n"
              f"pip install pandas openpyxl xlrd\nОшибка: {e}")
        sys.exit(1)

    root = tk.Tk()
    app = ArticleFinderGUI(root)
    root.mainloop()