"""
Приложение для поиска совпадающих артикулов в двух Excel файлах.
Архитектура построена по принципам SOLID.
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
        Регулярное выражение для поиска артикулов.
    """

    def __init__(self, pattern: str = r'[A-Za-zА-Яа-я0-9\-_]{3,}') -> None:
        self._pattern = re.compile(pattern)

    def extract(self, text: Optional[str]) -> Set[str]:
        """
        Извлекает уникальные артикулы из переданного текста.

        Parameters
        ----------
        text : str or None
            Текст для анализа.

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
        Загружает Excel файл в DataFrame.

        Returns
        -------
        pd.DataFrame
            Загруженная таблица данных.
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
        """Возвращает список названий колонок."""
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
        """
        if self._dataframe is None:
            raise ValueError("Данные не загружены.")
        if column_name not in self._dataframe.columns:
            raise KeyError(f"Колонка '{column_name}' не найдена.")
        return self._dataframe[column_name]

    def get_first_row(self) -> Optional[pd.Series]:
        """Возвращает первую строку данных для предпросмотра."""
        if self._dataframe is None or self._dataframe.empty:
            return None
        return self._dataframe.iloc[0]


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


# ... (начало файла без изменений до класса ArticleFinderGUI) ...

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
        self.root.geometry("900x750")
        self.root.resizable(True, True)

        self._extractor = ArticleExtractor()
        self._matcher = ArticleMatcher()
        self._loader1: Optional[ExcelDataLoader] = None
        self._loader2: Optional[ExcelDataLoader] = None

        # Хранилище выбранных колонок (по имени)
        self.selected_col1: Optional[str] = None
        self.selected_col2: Optional[str] = None
        
        # Хранилище оригинальных заголовков для сброса визуального выделения
        self.orig_headers1: List[str] = []
        self.orig_headers2: List[str] = []

        # Виджеты предпросмотра и статусов выбора
        self.preview_tree1: ttk.Treeview = None
        self.preview_tree2: ttk.Treeview = None
        self.lbl_sel1: ttk.Label = None
        self.lbl_sel2: ttk.Label = None

        self._setup_ui()

    def _setup_ui(self) -> None:
        """Инициализирует и размещает все элементы интерфейса."""
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=25)

        # --- Файл 1 ---
        frm1 = ttk.LabelFrame(self.root, text="Первый файл")
        frm1.pack(fill="x", padx=10, pady=5)

        btn_frame1 = ttk.Frame(frm1)
        btn_frame1.pack(fill="x", padx=5, pady=5)
        
        self.path1_var = tk.StringVar()
        ttk.Entry(btn_frame1, textvariable=self.path1_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        ttk.Button(btn_frame1, text="Выбрать", command=lambda: self._select_file(self.path1_var)).pack(
            side="left", padx=2
        )
        ttk.Button(btn_frame1, text="Загрузить", command=self._load_file1).pack(
            side="left", padx=2
        )

        # Статус выбора колонки 1
        self.lbl_sel1 = ttk.Label(frm1, text="Выбрана колонка: Не выбрана", foreground="grey")
        self.lbl_sel1.pack(anchor="w", padx=10, pady=(0, 5))

        # Предпросмотр данных 1
        self.preview_tree1 = self._create_preview_widget(frm1, file_num=1)

        # --- Файл 2 ---
        # ИЗМЕНЕНИЕ: Обновлён текст заголовка
        frm2 = ttk.LabelFrame(self.root, text="Второй файл (с которым сравниваем)")
        frm2.pack(fill="x", padx=10, pady=5)

        btn_frame2 = ttk.Frame(frm2)
        btn_frame2.pack(fill="x", padx=5, pady=5)
        
        self.path2_var = tk.StringVar()
        ttk.Entry(btn_frame2, textvariable=self.path2_var, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        ttk.Button(btn_frame2, text="Выбрать", command=lambda: self._select_file(self.path2_var)).pack(
            side="left", padx=2
        )
        ttk.Button(btn_frame2, text="Загрузить", command=self._load_file2).pack(
            side="left", padx=2
        )

        # Статус выбора колонки 2
        self.lbl_sel2 = ttk.Label(frm2, text="Выбрана колонка: Не выбрана", foreground="grey")
        self.lbl_sel2.pack(anchor="w", padx=10, pady=(0, 5))

        # Предпросмотр данных 2
        self.preview_tree2 = self._create_preview_widget(frm2, file_num=2)

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

    def _create_preview_widget(self, parent: ttk.Widget, file_num: int) -> ttk.Treeview:
        """
        Создаёт виджет предпросмотра данных с обработчиком клика по заголовкам.

        Parameters
        ----------
        parent : ttk.Widget
            Родительский фрейм.
        file_num : int
            Идентификатор файла (1 или 2) для привязки логики выбора.

        Returns
        -------
        ttk.Treeview
            Виджет предпросмотра.
        """
        preview_frame = ttk.Frame(parent)
        preview_frame.pack(fill="both", expand=False, padx=5, pady=(0, 5))  # Уменьшил pady
        
        ttk.Label(preview_frame, text="Предпросмотр (кликните по заголовку):").pack(anchor="w")
        
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill="x", expand=True)
        
        # Уменьшил высоту до 2 строк
        tree = ttk.Treeview(tree_frame, columns=(), show="headings", height=2)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=scroll_x.set)

        tree.bind('<Button-1>', lambda event: self._on_header_click(event, file_num, tree))

        tree.pack(side="left", fill="x", expand=True)
        scroll_x.pack(side="bottom", fill="x")
        return tree

    def _on_header_click(self, event: tk.Event, file_num: int, tree: ttk.Treeview) -> None:
        """
        Обработчик клика по заголовку колонки.

        Parameters
        ----------
        event : tk.Event
            Событие клика.
        file_num : int
            Номер файла.
        tree : ttk.Treeview
            Виджет, по которому кликнули.
        """
        region = tree.identify_region(event.x, event.y)
        if region != "heading":
            return

        col_id = tree.identify_column(event.x)
        col_index = int(col_id.lstrip("#")) - 1
        cols = tree["columns"]
        
        if col_index < 0 or col_index >= len(cols):
            return

        selected_col_name = cols[col_index]

        # Обновляем состояние
        if file_num == 1:
            self.selected_col1 = selected_col_name
            self._update_visual_selection(tree, cols, selected_col_name, self.orig_headers1, 1)
            self.lbl_sel1.config(text=f"Выбрана колонка: {selected_col_name}", foreground="black")
        else:
            self.selected_col2 = selected_col_name
            self._update_visual_selection(tree, cols, selected_col_name, self.orig_headers2, 2)
            self.lbl_sel2.config(text=f"Выбрана колонка: {selected_col_name}", foreground="black")

    def _update_visual_selection(self, tree: ttk.Treeview, cols: tuple, 
                                 selected_col: str, orig_headers: list, file_num: int) -> None:
        """
        Обновляет текст заголовков: сбрасывает выделение со всех и ставит маркер на выбранную.

        Parameters
        ----------
        tree : ttk.Treeview
            Виджет дерева.
        cols : tuple
            Кортеж идентификаторов колонок.
        selected_col : str
            Имя выбранной колонки.
        orig_headers : list
            Список оригинальных заголовков.
        file_num : int
            Номер файла (для сохранения оригиналов).
        """
        if not orig_headers:
            orig_headers.clear()
            for col in cols:
                orig_headers.append(tree.heading(col, option="text"))

        for col in cols:
            tree.heading(col, text=orig_headers[cols.index(col)])

        tree.heading(selected_col, text=f"▶ {orig_headers[cols.index(selected_col)]}")

    def _update_preview(self, tree: ttk.Treeview, first_row: pd.Series, file_num: int) -> None:
        """
        Обновляет содержимое виджета предпросмотра данными.

        Parameters
        ----------
        tree : ttk.Treeview
            Виджет для обновления.
        first_row : pd.Series
            Данные первой строки.
        file_num : int
            Номер файла (для сброса выбора).
        """
        tree.delete(*tree.get_children())
        tree["columns"] = ()

        if file_num == 1:
            self.selected_col1 = None
            self.orig_headers1.clear()
            self.lbl_sel1.config(text="Выбрана колонка: Не выбрана", foreground="grey")
        else:
            self.selected_col2 = None
            self.orig_headers2.clear()
            self.lbl_sel2.config(text="Выбрана колонка: Не выбрана", foreground="grey")

        if first_row is None:
            return

        cols = first_row.index.astype(str).tolist()
        tree["columns"] = cols

        # Расчёт ширины колонок с ограничением
        max_width = 150  # Максимальная ширина колонки
        min_width = 80   # Минимальная ширина
        
        for col in cols:
            header_text = str(col)
            cell_value = str(first_row[col]) if pd.notna(first_row[col]) else ""
            
            # Усекаем длинные значения (добавляем "...")
            if len(cell_value) > 40:
                cell_value = cell_value[:37] + "..."
            
            # Расчитываем ширину на основе длины заголовка и значения
            max_len = max(len(header_text), len(cell_value))
            calculated_width = max(min_width, min(max_len * 7 + 10, max_width))
            
            tree.heading(col, text=header_text)
            tree.column(col, width=calculated_width, anchor="center", stretch=False)

        # Подготавливаем значения (с усечением)
        values = []
        for val in first_row.values:
            str_val = str(val) if pd.notna(val) else ""
            if len(str_val) > 40:
                str_val = str_val[:37] + "..."
            values.append(str_val)
        
        tree.insert("", "end", values=values)

    def _select_file(self, string_var: tk.StringVar) -> None:
        """Открывает диалог выбора файла."""
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if filepath:
            string_var.set(filepath)

    def _load_file1(self) -> None:
        """Загружает первый Excel файл."""
        self._load_file(self.path1_var.get(), 1)

    def _load_file2(self) -> None:
        """Загружает второй Excel файл."""
        self._load_file(self.path2_var.get(), 2)

    def _load_file(self, path: str, file_num: int) -> None:
        """
        Общая логика загрузки Excel файла.

        Parameters
        ----------
        path : str
            Путь к файлу.
        file_num : int
            Номер файла (1 или 2).
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
            else:
                self._loader2 = loader
            
            tree = self.preview_tree1 if file_num == 1 else self.preview_tree2
            self._update_preview(tree, loader.get_first_row(), file_num)

            messagebox.showinfo("Успех", f"Файл {file_num} загружен. Кликните по заголовку для выбора.")
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    def _find_matches(self) -> None:
        """Ищет совпадения в выбранных колонках."""
        if not self._loader1 or not self._loader2:
            messagebox.showwarning("Внимание", "Загрузите оба файла перед поиском.")
            return

        if not self.selected_col1 or not self.selected_col2:
            messagebox.showwarning("Внимание", "Выберите колонки для анализа кликом по заголовку в таблице.")
            return

        try:
            self._update_status("Анализ первого файла...")
            self.root.update()
            articles1 = set()
            for val in self._loader1.get_column_data(self.selected_col1).dropna():
                articles1.update(self._extractor.extract(str(val)))

            self._update_status("Анализ второго файла...")
            self.root.update()
            articles2 = set()
            for val in self._loader2.get_column_data(self.selected_col2).dropna():
                articles2.update(self._extractor.extract(str(val)))

            self._update_status("Поиск совпадений...")
            self.root.update()
            common = self._matcher.find_common(articles1, articles2)

            self._display_results(common)
        except Exception as e:
            messagebox.showerror("Ошибка поиска", str(e))
            self._update_status("Произошла ошибка. Проверьте данные.")

    def _update_status(self, text: str) -> None:
        """Обновляет поле статуса."""
        self.res_text.config(state="normal")
        self.res_text.delete("1.0", tk.END)
        self.res_text.insert("1.0", text)
        self.res_text.config(state="disabled")
        self.root.update()

    def _display_results(self, matches: Set[str]) -> None:
        """Выводит результаты поиска."""
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
    try:
        import pandas as pd
        import openpyxl
        import xlrd
    except ImportError as e:
        print(f"Установите библиотеки: pip install pandas openpyxl xlrd")
        sys.exit(1)

    root = tk.Tk()
    app = ArticleFinderGUI(root)
    root.mainloop()