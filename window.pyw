"""
Приложение для поиска строк по артикулам в Excel файлах.
Архитектура: SOLID, NumPy docstring style.
"""

import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import Optional, Set, List, Dict

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


class TemplateManager:
    """Менеджер управления шаблонами регулярных выражений."""

    def __init__(self, templates_dir: str = "templates") -> None:
        self._templates_dir = templates_dir
        self._ensure_templates_dir()
        self._current_pattern: Optional[str] = None
        self._current_template_name: Optional[str] = None

    def _ensure_templates_dir(self) -> None:
        """Создаёт каталог для шаблонов, если он не существует."""
        if not os.path.exists(self._templates_dir):
            os.makedirs(self._templates_dir)
            example_path = os.path.join(self._templates_dir, "example_patterns.txt")
            if not os.path.exists(example_path):
                with open(example_path, 'w', encoding='utf-8') as f:
                    f.write("# Шаблон для артикулов обоев\n")
                    f.write(r"\b(?:[A-Z]{0,3}\d{3,6}(?:[-–]\d{1,3})?|\d{4,6}(?:[-–][A-Z0-9]{1,3})?)\b")

    def get_templates_dir(self) -> str:
        return self._templates_dir

    def load_template(self, template_path: str) -> Optional[str]:
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            for line in lines:
                line = line.strip()
                if line and not line.startswith('#'):
                    self._current_pattern = line
                    self._current_template_name = os.path.basename(template_path)
                    return line
            messagebox.showwarning("Предупреждение", "Файл шаблона пуст или содержит только комментарии.")
            return None
        except Exception as e:
            messagebox.showerror("Ошибка загрузки шаблона", str(e))
            return None

    def get_current_pattern(self) -> Optional[str]:
        return self._current_pattern

    def get_current_template_name(self) -> Optional[str]:
        return self._current_template_name

    def reset_to_default(self) -> None:
        self._current_pattern = r'[A-Za-zА-Яа-я0-9\-_]{3,}'
        self._current_template_name = "По умолчанию"


class ExcelDataLoader:
    """Класс для загрузки Excel файлов."""

    def __init__(self, file_path: str) -> None:
        self._file_path = file_path
        self._dataframe: Optional[pd.DataFrame] = None

    def load(self) -> pd.DataFrame:
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
        # Заполняем NaN пустыми строками для удобства работы с текстом
        self._dataframe = self._dataframe.fillna('')
        return self._dataframe

    def get_columns(self) -> List[str]:
        if self._dataframe is None:
            return []
        return self._dataframe.columns.astype(str).tolist()

    def get_column_data(self, column_name: str) -> pd.Series:
        if self._dataframe is None:
            raise ValueError("Данные не загружены.")
        if column_name not in self._dataframe.columns:
            raise KeyError(f"Колонка '{column_name}' не найдена.")
        return self._dataframe[column_name]

    def get_first_row(self) -> Optional[pd.Series]:
        if self._dataframe is None or self._dataframe.empty:
            return None
        return self._dataframe.iloc[0]


class ArticleSearchEngine:
    """
    Движок поиска строк по списку артикулов.
    """

    def __init__(self, extractor: ArticleExtractor) -> None:
        self._extractor = extractor

    def search(
        self,
        source_df: pd.DataFrame,
        source_col: str,
        target_df: pd.DataFrame,
        target_col: str,
        output_cols: List[str]
    ) -> pd.DataFrame:
        """
        Ищет строки в target_df, содержащие артикулы из source_df.

        Parameters
        ----------
        source_df : pd.DataFrame
            DataFrame со списком эталонных артикулов.
        source_col : str
            Колонка в source_df с чистыми артикулами.
        target_df : pd.DataFrame
            DataFrame, в котором производится поиск.
        target_col : str
            Колонка в target_df, где ищутся совпадения.
        output_cols : List[str]
            Список колонок для включения в результат.

        Returns
        -------
        pd.DataFrame
            Отфильтрованный DataFrame с выбранными колонками.
        """
        # 1. Собираем множество эталонных артикулов
        source_articles = set(source_df[source_col].astype(str).str.strip().unique())
        source_articles.discard('')
        source_articles.discard('nan')

        if not source_articles:
            raise ValueError("В выбранной колонке первого файла не найдено артикулов.")

        # 2. Фильтруем второй файл
        mask = target_df[target_col].apply(
            lambda x: bool(self._extractor.extract(str(x)).intersection(source_articles))
        )

        filtered_df = target_df.loc[mask].copy()

        if filtered_df.empty:
            return pd.DataFrame()

        # 3. Выбираем нужные колонки
        valid_cols = [col for col in output_cols if col in filtered_df.columns]
        if not valid_cols:
            valid_cols = filtered_df.columns.tolist()

        result_df = filtered_df[valid_cols]
        return result_df


class ArticleFinderGUI:
    """Графический интерфейс приложения."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Поиск строк по артикулам")
        self.root.geometry("950x700")
        self.root.resizable(True, True)

        self._template_manager = TemplateManager()
        self._extractor = ArticleExtractor(self._template_manager.get_current_pattern() or r'[A-Za-zА-Яа-я0-9\-_]{3,}')
        self._search_engine = ArticleSearchEngine(self._extractor)
        
        self._loader1: Optional[ExcelDataLoader] = None
        self._loader2: Optional[ExcelDataLoader] = None

        # Состояние выбора колонок (храним внутренние имена колонок pandas)
        self.selected_src_col: Optional[str] = None
        self.selected_tgt_col: Optional[str] = None
        
        # Выбранные колонки для вывода (внутренние имена)
        self.selected_output_cols: Set[str] = set()

        # Виджеты
        self.lbl_template: ttk.Label = None
        
        # Метки статуса выбора
        self.lbl_src_status: ttk.Label = None
        self.lbl_tgt_status: ttk.Label = None
        
        # Деревья предпросмотра
        self.preview_tree1: ttk.Treeview = None
        self.preview_tree2: ttk.Treeview = None
        
        # Фрейм с чекбоксами
        self.checkboxes_frame: ttk.Frame = None
        self.col_checkboxes: Dict[str, ttk.Checkbutton] = {}

        self._setup_ui()
        self._create_context_menu()

    def _format_column_name(self, col_name: str, index: int) -> str:
        """
        Форматирует имя колонки для отображения пользователю.
        Если имя пустое или Unnamed, возвращает 'Колонка N'.
        """
        if not col_name or col_name.startswith('Unnamed'):
            return f"Колонка {index + 1}"
        return col_name

    def _setup_ui(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        # --- Панель шаблонов ---
        frm_template = ttk.Frame(self.root)
        frm_template.pack(fill="x", padx=10, pady=5)
        ttk.Label(frm_template, text="Шаблон поиска:").pack(side="left", padx=(0, 5))
        self.lbl_template = ttk.Label(frm_template, text="По умолчанию", foreground="blue")
        self.lbl_template.pack(side="left")
        ttk.Button(frm_template, text="Загрузить шаблон", command=self._load_template_via_dialog).pack(side="left", padx=(10, 0))
        ttk.Button(frm_template, text="Сбросить", command=self._reset_template).pack(side="left", padx=(5, 0))

        # --- Файл 1 (Источник артикулов) ---
        frm1 = ttk.LabelFrame(self.root, text="1. Файл со списком артикулов (Эталон)")
        frm1.pack(fill="x", padx=10, pady=5)
        
        btn_frame1 = ttk.Frame(frm1)
        btn_frame1.pack(fill="x", padx=5, pady=5)
        self.path1_var = tk.StringVar()
        ttk.Entry(btn_frame1, textvariable=self.path1_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame1, text="Выбрать", command=lambda: self._select_file(self.path1_var)).pack(side="left", padx=2)
        ttk.Button(btn_frame1, text="Загрузить", command=self._load_file1).pack(side="left", padx=2)

        # Статус выбора колонки 1
        self.lbl_src_status = ttk.Label(frm1, text="Выбрана колонка: Не выбрана (кликните по заголовку)", foreground="grey")
        self.lbl_src_status.pack(anchor="w", padx=10, pady=(0, 5))

        # Предпросмотр Файла 1
        self.preview_tree1 = self._create_preview_widget(frm1, file_num=1)

        # --- Файл 2 (Целевая база) ---
        frm2 = ttk.LabelFrame(self.root, text="2. Файл для поиска (Где искать)")
        frm2.pack(fill="x", padx=10, pady=5)
        
        btn_frame2 = ttk.Frame(frm2)
        btn_frame2.pack(fill="x", padx=5, pady=5)
        self.path2_var = tk.StringVar()
        ttk.Entry(btn_frame2, textvariable=self.path2_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame2, text="Выбрать", command=lambda: self._select_file(self.path2_var)).pack(side="left", padx=2)
        ttk.Button(btn_frame2, text="Загрузить", command=self._load_file2).pack(side="left", padx=2)

        # Статус выбора колонки поиска
        self.lbl_tgt_status = ttk.Label(frm2, text="Колонка для поиска: Не выбрана (кликните по заголовку)", foreground="grey")
        self.lbl_tgt_status.pack(anchor="w", padx=10, pady=(0, 5))

        # Предпросмотр Файла 2
        self.preview_tree2 = self._create_preview_widget(frm2, file_num=2)

        # Панель выбора колонок для вывода
        self.checkboxes_frame = ttk.LabelFrame(frm2, text="Колонки для вывода (отметьте нужные)")
        self.checkboxes_frame.pack(fill="both", expand=False, padx=5, pady=5)
        # Сюда будут добавляться чекбоксы динамически

        # --- Кнопка поиска ---
        ttk.Button(self.root, text="Найти совпадения", command=self._find_matches).pack(pady=10)

        # --- Результаты ---
        frm_res = ttk.LabelFrame(self.root, text="Результаты поиска")
        frm_res.pack(fill="both", expand=True, padx=10, pady=5)

        tree_frame = ttk.Frame(frm_res)
        tree_frame.pack(fill="both", expand=True)
        
        self.result_tree = ttk.Treeview(tree_frame, show="headings")
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.result_tree.xview)
        
        self.result_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.result_tree.pack(side="left", fill="both", expand=True)

    def _create_preview_widget(self, parent: ttk.Widget, file_num: int) -> ttk.Treeview:
        preview_frame = ttk.Frame(parent)
        preview_frame.pack(fill="both", expand=False, padx=5, pady=(0, 5))
        ttk.Label(preview_frame, text="Предпросмотр (кликните по заголовку для выбора):").pack(anchor="w")
        
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill="x", expand=True)
        
        tree = ttk.Treeview(tree_frame, columns=(), show="headings", height=2)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=scroll_x.set)
        
        # Привязка клика
        tree.bind('<Button-1>', lambda event: self._on_header_click(event, file_num, tree))
        
        tree.pack(side="left", fill="x", expand=True)
        scroll_x.pack(side="bottom", fill="x")
        return tree

    def _create_context_menu(self) -> None:
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Открыть каталог шаблонов", command=self._open_templates_dir)
        self.context_menu.add_command(label="Загрузить шаблон...", command=self._load_template_via_dialog)
        self.context_menu.add_command(label="Сбросить к стандартному", command=self._reset_template)
        self.root.bind("<Button-3>", self._show_context_menu)

    def _show_context_menu(self, event: tk.Event) -> None:
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _open_templates_dir(self) -> None:
        templates_dir = self._template_manager.get_templates_dir()
        if os.name == 'nt':
            os.startfile(templates_dir)
        elif os.name == 'posix':
            import subprocess
            subprocess.call(['xdg-open', templates_dir])

    def _load_template_via_dialog(self) -> None:
        templates_dir = self._template_manager.get_templates_dir()
        filepath = filedialog.askopenfilename(initialdir=templates_dir, filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")])
        if filepath:
            self._load_template(filepath)

    def _load_template(self, template_path: str) -> None:
        pattern = self._template_manager.load_template(template_path)
        if pattern:
            try:
                re.compile(pattern)
                self._extractor = ArticleExtractor(pattern)
                self._search_engine = ArticleSearchEngine(self._extractor)
                self.lbl_template.config(text=f"{self._template_manager.get_current_template_name()}", foreground="green")
                messagebox.showinfo("Успех", f"Шаблон '{self._template_manager.get_current_template_name()}' загружен.")
            except re.error as e:
                messagebox.showerror("Ошибка шаблона", f"Невалидное регулярное выражение:\n{e}")
                self._reset_template()

    def _reset_template(self) -> None:
        self._template_manager.reset_to_default()
        self._extractor = ArticleExtractor(self._template_manager.get_current_pattern())
        self._search_engine = ArticleSearchEngine(self._extractor)
        self.lbl_template.config(text="По умолчанию", foreground="blue")

    def _select_file(self, string_var: tk.StringVar) -> None:
        filepath = filedialog.askopenfilename(filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")])
        if filepath:
            string_var.set(filepath)

    def _load_file1(self) -> None:
        self._load_file(self.path1_var.get(), 1)

    def _load_file2(self) -> None:
        self._load_file(self.path2_var.get(), 2)

    def _load_file(self, path: str, file_num: int) -> None:
        if not path:
            messagebox.showwarning("Внимание", "Сначала выберите файл.")
            return

        try:
            loader = ExcelDataLoader(path)
            loader.load()
            
            if file_num == 1:
                self._loader1 = loader
                self.selected_src_col = None
                self.lbl_src_status.config(text="Выбрана колонка: Не выбрана (кликните по заголовку)", foreground="grey")
                self._update_preview(self.preview_tree1, loader.get_first_row(), 1)
            else:
                self._loader2 = loader
                self.selected_tgt_col = None
                self.selected_output_cols.clear()
                self.lbl_tgt_status.config(text="Колонка для поиска: Не выбрана (кликните по заголовку)", foreground="grey")
                self._update_preview(self.preview_tree2, loader.get_first_row(), 2)
                self._update_checkboxes(loader.get_columns())

            messagebox.showinfo("Успех", f"Файл {file_num} загружен.")
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    def _on_header_click(self, event: tk.Event, file_num: int, tree: ttk.Treeview) -> None:
        region = tree.identify_region(event.x, event.y)
        if region != "heading":
            return

        col_id = tree.identify_column(event.x)
        col_index = int(col_id.lstrip("#")) - 1
        cols = tree["columns"]
        
        if col_index < 0 or col_index >= len(cols):
            return

        # col_id в Treeview - это внутренний идентификатор (например '#1'), 
        # но нам нужно реальное имя колонки из списка tree["columns"]
        # tree["columns"] содержит tuple идентификаторов колонок.
        # В нашем случае при создании tree["columns"] = cols, где cols - это имена из DataFrame.
        
        selected_col_name = cols[col_index]

        if file_num == 1:
            self.selected_src_col = selected_col_name
            display_name = self._format_column_name(selected_col_name, col_index)
            self.lbl_src_status.config(text=f"Выбрана колонка: {display_name}", foreground="black")
            # Визуально выделяем заголовок
            self._highlight_header(tree, cols, selected_col_name)
        else:
            self.selected_tgt_col = selected_col_name
            display_name = self._format_column_name(selected_col_name, col_index)
            self.lbl_tgt_status.config(text=f"Колонка для поиска: {display_name}", foreground="black")
            self._highlight_header(tree, cols, selected_col_name)

    def _highlight_header(self, tree: ttk.Treeview, cols: tuple, selected_col: str) -> None:
        # Сброс стиля для всех колонок
        for col in cols:
            tree.heading(col, text=self._format_column_name(col, list(cols).index(col)))
        
        # Установка выделения
        idx = list(cols).index(selected_col)
        tree.heading(selected_col, text=f"▶ {self._format_column_name(selected_col, idx)}")

    def _update_preview(self, tree: ttk.Treeview, first_row: pd.Series, file_num: int) -> None:
        tree.delete(*tree.get_children())
        tree["columns"] = ()
        if first_row is None:
            return
        
        cols = first_row.index.astype(str).tolist()
        tree["columns"] = cols
        
        for i, col in enumerate(cols):
            tree.heading(col, text=self._format_column_name(col, i))
            tree.column(col, width=120, anchor="center")
            
        tree.insert("", "end", values=[str(val) for val in first_row.values])

    def _update_checkboxes(self, columns: List[str]) -> None:
        # Очистка старых чекбоксов
        for widget in self.checkboxes_frame.winfo_children():
            widget.destroy()
        self.col_checkboxes.clear()

        # Кнопки "Все" / "Ничего"
        btn_frame = ttk.Frame(self.checkboxes_frame)
        btn_frame.pack(fill="x", padx=5, pady=2)
        ttk.Button(btn_frame, text="Все", command=lambda: self._toggle_all(True)).pack(side="left", padx=2)
        ttk.Button(btn_frame, text="Ничего", command=lambda: self._toggle_all(False)).pack(side="left", padx=2)

        # Создание чекбоксов
        for i, col in enumerate(columns):
            display_name = self._format_column_name(col, i)
            var = tk.BooleanVar(value=True)  # По умолчанию все выбраны
            cb = ttk.Checkbutton(self.checkboxes_frame, text=display_name, variable=var)
            cb.pack(anchor="w", padx=20, pady=1)
            self.col_checkboxes[col] = var

    def _toggle_all(self, value: bool) -> None:
        for var in self.col_checkboxes.values():
            var.set(value)

    def _find_matches(self) -> None:
        if not self._loader1 or not self._loader2:
            messagebox.showwarning("Внимание", "Загрузите оба файла перед поиском.")
            return
        if not self.selected_src_col or not self.selected_tgt_col:
            messagebox.showwarning("Внимание", "Выберите колонки для анализа (кликните по заголовкам таблиц).")
            return

        # Сбор выбранных колонок для вывода
        output_cols = [col for col, var in self.col_checkboxes.items() if var.get()]
        if not output_cols:
            messagebox.showwarning("Внимание", "Выберите хотя бы одну колонку для вывода.")
            return

        try:
            self.root.config(cursor="wait")
            self.root.update()
            
            result_df = self._search_engine.search(
                self._loader1._dataframe,
                self.selected_src_col,
                self._loader2._dataframe,
                self.selected_tgt_col,
                output_cols
            )
            
            self._display_results(result_df)
            
            if result_df.empty:
                messagebox.showinfo("Результат", "Совпадений не найдено.")
            else:
                messagebox.showinfo("Результат", f"Найдено строк: {len(result_df)}")
                
        except Exception as e:
            messagebox.showerror("Ошибка поиска", str(e))
        finally:
            self.root.config(cursor="")

    def _display_results(self, df: pd.DataFrame) -> None:
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        if df.empty:
            self.result_tree["columns"] = ()
            return

        cols = df.columns.astype(str).tolist()
        self.result_tree["columns"] = cols
        
        for i, col in enumerate(cols):
            display_name = self._format_column_name(col, i)
            self.result_tree.heading(col, text=display_name)
            max_len = df[col].astype(str).map(len).max()
            width = max(50, min(300, max_len * 7 + 10))
            self.result_tree.column(col, width=width, anchor="w")

        data = df.values.tolist()
        for row in data:
            str_row = [str(val) for val in row]
            self.result_tree.insert("", "end", values=str_row)


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