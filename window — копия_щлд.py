"""
Приложение для поиска строк по артикулам в Excel файлах.
Обновленная версия с выбором колонок прямо в предпросмотре.
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
    """
    def __init__(self, pattern: str = r'[A-Za-zА-Яа-я0-9\-_]{3,}') -> None:
        self._pattern = re.compile(pattern)

    def extract(self, text: Optional[str]) -> Set[str]:
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
            messagebox.showwarning("Предупреждение", "Файл шаблона пуст.")
            return None
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
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
        if ext == ".xlsx": engine = "openpyxl"
        elif ext == ".xls": engine = "xlrd"
        else: raise ValueError("Поддерживаются только форматы .xls и .xlsx")

        self._dataframe = pd.read_excel(self._file_path, engine=engine)
        self._dataframe = self._dataframe.fillna('')
        return self._dataframe

    def get_columns(self) -> List[str]:
        if self._dataframe is None: return []
        return self._dataframe.columns.astype(str).tolist()

    def get_column_data(self, column_name: str) -> pd.Series:
        if self._dataframe is None: raise ValueError("Данные не загружены.")
        if column_name not in self._dataframe.columns: raise KeyError(f"Колонка '{column_name}' не найдена.")
        return self._dataframe[column_name]

    def get_first_row(self) -> Optional[pd.Series]:
        if self._dataframe is None or self._dataframe.empty: return None
        return self._dataframe.iloc[0]


class ArticleSearchEngine:
    """Движок поиска строк по списку артикулов."""
    def __init__(self, extractor: ArticleExtractor) -> None:
        self._extractor = extractor

    def search(self, source_df: pd.DataFrame, source_col: str, 
               target_df: pd.DataFrame, target_col: str, output_cols: List[str]) -> pd.DataFrame:
        
        # Собираем эталонные артикулы
        source_articles = set(source_df[source_col].astype(str).str.strip().unique())
        source_articles.discard('')
        source_articles.discard('nan')
        if not source_articles: raise ValueError("В колонке артикулов пусто.")

        # Фильтруем второй файл
        mask = target_df[target_col].apply(
            lambda x: bool(self._extractor.extract(str(x)).intersection(source_articles))
        )
        filtered_df = target_df.loc[mask].copy()

        if filtered_df.empty: return pd.DataFrame()

        # Выбираем нужные колонки
        valid_cols = [col for col in output_cols if col in filtered_df.columns]
        if not valid_cols: valid_cols = filtered_df.columns.tolist()
        
        return filtered_df[valid_cols]


class ArticleFinderGUI:
    """Графический интерфейс приложения."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Поиск строк по артикулам")
        self.root.geometry("900x650")
        self.root.resizable(True, True)

        self._template_manager = TemplateManager()
        self._extractor = ArticleExtractor(self._template_manager.get_current_pattern() or r'[A-Za-zА-Яа-я0-9\-_]{3,}')
        self._search_engine = ArticleSearchEngine(self._extractor)
        
        self._loader1: Optional[ExcelDataLoader] = None
        self._loader2: Optional[ExcelDataLoader] = None

        # Состояние выбора колонок
        self.selected_src_col: Optional[str] = None  # Колонка с артикулами (Файл 1)
        self.selected_tgt_col: Optional[str] = None  # Колонка для поиска (Файл 2)
        self.selected_output_cols: Set[str] = set()  # Колонки для вывода (Файл 2)

        # Виджеты
        self.lbl_template: ttk.Label = None
        self.preview_tree1: ttk.Treeview = None
        self.preview_tree2: ttk.Treeview = None
        
        # Метки статуса
        self.lbl_src_status: ttk.Label = None
        self.lbl_tgt_status: ttk.Label = None
        self.lbl_out_status: ttk.Label = None

        self._setup_ui()
        self._create_context_menu()

    def _format_column_name(self, col_name: str, index: int) -> str:
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

        # --- Файл 1 (Источник) ---
        frm1 = ttk.LabelFrame(self.root, text="1. Файл со списком артикулов (Эталон)")
        frm1.pack(fill="x", padx=10, pady=5)
        
        btn_frame1 = ttk.Frame(frm1)
        btn_frame1.pack(fill="x", padx=5, pady=5)
        self.path1_var = tk.StringVar()
        ttk.Entry(btn_frame1, textvariable=self.path1_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame1, text="Выбрать", command=lambda: self._select_file(self.path1_var)).pack(side="left", padx=2)
        ttk.Button(btn_frame1, text="Загрузить", command=self._load_file1).pack(side="left", padx=2)

        self.lbl_src_status = ttk.Label(frm1, text="Выбрана колонка с артикулами: Не выбрана (ЛКМ по заголовку)", foreground="grey")
        self.lbl_src_status.pack(anchor="w", padx=10, pady=(0, 5))
        self.preview_tree1 = self._create_preview_widget(frm1, file_num=1)

        # --- Файл 2 (Целевой) ---
        frm2 = ttk.LabelFrame(self.root, text="2. Файл для поиска (Где искать)")
        frm2.pack(fill="x", padx=10, pady=5)
        
        btn_frame2 = ttk.Frame(frm2)
        btn_frame2.pack(fill="x", padx=5, pady=5)
        self.path2_var = tk.StringVar()
        ttk.Entry(btn_frame2, textvariable=self.path2_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame2, text="Выбрать", command=lambda: self._select_file(self.path2_var)).pack(side="left", padx=2)
        ttk.Button(btn_frame2, text="Загрузить", command=self._load_file2).pack(side="left", padx=2)

        # Статусы для второго файла
        status_frame2 = ttk.Frame(frm2)
        status_frame2.pack(fill="x", padx=10, pady=(0, 5))
        
        self.lbl_tgt_status = ttk.Label(status_frame2, text="Колонка поиска: Не выбрана (ЛКМ)", foreground="grey")
        self.lbl_tgt_status.pack(side="left", padx=(0, 20))
        
        self.lbl_out_status = ttk.Label(status_frame2, text="Колонки вывода: Не выбраны (ПКМ)", foreground="grey")
        self.lbl_out_status.pack(side="left")

        self.preview_tree2 = self._create_preview_widget(frm2, file_num=2)

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
        ttk.Label(preview_frame, text="Предпросмотр:").pack(anchor="w")
        
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill="x", expand=True)
        
        tree = ttk.Treeview(tree_frame, columns=(), show="headings", height=2)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=scroll_x.set)
        
        # Привязки событий
        tree.bind('<Button-1>', lambda event: self._on_header_click(event, file_num, tree)) # ЛКМ - Выбор поиска
        tree.bind('<Button-3>', lambda event: self._on_header_rclick(event, file_num, tree)) # ПКМ - Выбор вывода
        
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
        if os.name == 'nt': os.startfile(templates_dir)
        elif os.name == 'posix':
            import subprocess
            subprocess.call(['xdg-open', templates_dir])

    def _load_template_via_dialog(self) -> None:
        templates_dir = self._template_manager.get_templates_dir()
        filepath = filedialog.askopenfilename(initialdir=templates_dir, filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")])
        if filepath: self._load_template(filepath)

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
                messagebox.showerror("Ошибка шаблона", f"Невалидное выражение:\n{e}")
                self._reset_template()

    def _reset_template(self) -> None:
        self._template_manager.reset_to_default()
        self._extractor = ArticleExtractor(self._template_manager.get_current_pattern())
        self._search_engine = ArticleSearchEngine(self._extractor)
        self.lbl_template.config(text="По умолчанию", foreground="blue")

    def _select_file(self, string_var: tk.StringVar) -> None:
        filepath = filedialog.askopenfilename(filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")])
        if filepath: string_var.set(filepath)

    def _load_file1(self) -> None: self._load_file(self.path1_var.get(), 1)
    def _load_file2(self) -> None: self._load_file(self.path2_var.get(), 2)

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
                self.lbl_src_status.config(text="Выбрана колонка с артикулами: Не выбрана (ЛКМ по заголовку)", foreground="grey")
                self._update_preview(self.preview_tree1, loader.get_first_row(), 1)
            else:
                self._loader2 = loader
                self.selected_tgt_col = None
                self.selected_output_cols.clear()
                self.lbl_tgt_status.config(text="Колонка поиска: Не выбрана (ЛКМ)", foreground="grey")
                self.lbl_out_status.config(text="Колонки вывода: Не выбраны (ПКМ)", foreground="grey")
                self._update_preview(self.preview_tree2, loader.get_first_row(), 2)

            messagebox.showinfo("Успех", f"Файл {file_num} загружен.")
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    # --- ОБРАБОТЧИКИ КЛИКОВ ---

    def _on_header_click(self, event: tk.Event, file_num: int, tree: ttk.Treeview) -> None:
        """ЛКМ: Выбор колонки для поиска (для файла 2) или выбор артикулов (для файла 1)."""
        region = tree.identify_region(event.x, event.y)
        if region != "heading": return

        col_id = tree.identify_column(event.x)
        col_index = int(col_id.lstrip("#")) - 1
        cols = tree["columns"]
        if col_index < 0 or col_index >= len(cols): return
        
        selected_col_name = cols[col_index]

        if file_num == 1:
            self.selected_src_col = selected_col_name
            display_name = self._format_column_name(selected_col_name, col_index)
            self.lbl_src_status.config(text=f"Выбрана колонка с артикулами: {display_name}", foreground="black")
            self._highlight_header_single(tree, cols, selected_col_name)
        else:
            self.selected_tgt_col = selected_col_name
            display_name = self._format_column_name(selected_col_name, col_index)
            self.lbl_tgt_status.config(text=f"Колонка поиска: {display_name}", foreground="black")
            self._highlight_header_search(tree, cols, selected_col_name)

    def _on_header_rclick(self, event: tk.Event, file_num: int, tree: ttk.Treeview) -> None:
        """ПКМ: Добавление/Удаление колонки из списка вывода (только для файла 2)."""
        if file_num != 2: return

        region = tree.identify_region(event.x, event.y)
        if region != "heading": return

        col_id = tree.identify_column(event.x)
        col_index = int(col_id.lstrip("#")) - 1
        cols = tree["columns"]
        if col_index < 0 or col_index >= len(cols): return

        selected_col_name = cols[col_index]

        if selected_col_name in self.selected_output_cols:
            self.selected_output_cols.remove(selected_col_name)
        else:
            self.selected_output_cols.add(selected_col_name)

        # Обновляем статус
        if self.selected_output_cols:
            names = [self._format_column_name(c, list(cols).index(c)) for c in self.selected_output_cols]
            self.lbl_out_status.config(text=f"Колонки вывода: {', '.join(names)}", foreground="black")
        else:
            self.lbl_out_status.config(text="Колонки вывода: Не выбраны (ПКМ)", foreground="grey")

        self._highlight_header_output(tree, cols)

    # --- ВИЗУАЛИЗАЦИЯ ---

    def _highlight_header_single(self, tree: ttk.Treeview, cols: tuple, selected_col: str) -> None:
        """Подсветка для Файла 1 (один выбор)"""
        for col in cols:
            idx = list(cols).index(col)
            tree.heading(col, text=self._format_column_name(col, idx))
        
        idx = list(cols).index(selected_col)
        tree.heading(selected_col, text=f"▶ {self._format_column_name(selected_col, idx)}")

    def _highlight_header_search(self, tree: ttk.Treeview, cols: tuple, selected_col: str) -> None:
        """Подсветка для Файла 2 (Колонка поиска)"""
        # Сначала сбрасываем все заголовки, учитывая выбор вывода
        for col in cols:
            idx = list(cols).index(col)
            base_text = self._format_column_name(col, idx)
            if col in self.selected_output_cols:
                tree.heading(col, text=f"★ {base_text}")
            else:
                tree.heading(col, text=base_text)
        
        # Устанавливаем поиск
        idx = list(cols).index(selected_col)
        base_text = self._format_column_name(selected_col, idx)
        tree.heading(selected_col, text=f"▶ {base_text}")

    def _highlight_header_output(self, tree: ttk.Treeview, cols: tuple) -> None:
        """Обновление заголовков для Файла 2 (Колонки вывода)"""
        # Нужно сохранить текущий выбор поиска, чтобы не сбросить его визуально
        current_search = self.selected_tgt_col
        
        for col in cols:
            idx = list(cols).index(col)
            base_text = self._format_column_name(col, idx)
            
            # Логика формирования текста заголовка
            is_output = col in self.selected_output_cols
            is_search = col == current_search
            
            if is_search and is_output:
                tree.heading(col, text=f"▶ ★ {base_text}")
            elif is_search:
                tree.heading(col, text=f"▶ {base_text}")
            elif is_output:
                tree.heading(col, text=f"★ {base_text}")
            else:
                tree.heading(col, text=base_text)

    def _update_preview(self, tree: ttk.Treeview, first_row: pd.Series, file_num: int) -> None:
        tree.delete(*tree.get_children())
        tree["columns"] = ()
        if first_row is None: return
        
        cols = first_row.index.astype(str).tolist()
        tree["columns"] = cols
        
        for i, col in enumerate(cols):
            tree.heading(col, text=self._format_column_name(col, i))
            tree.column(col, width=120, anchor="center")
            
        tree.insert("", "end", values=[str(val) for val in first_row.values])

    def _find_matches(self) -> None:
        if not self._loader1 or not self._loader2:
            messagebox.showwarning("Внимание", "Загрузите оба файла перед поиском.")
            return
        if not self.selected_src_col or not self.selected_tgt_col:
            messagebox.showwarning("Внимание", "Выберите колонки (ЛКМ по заголовкам).")
            return
        if not self.selected_output_cols:
            messagebox.showwarning("Внимание", "Выберите хотя бы одну колонку для вывода (ПКМ по заголовкам).")
            return

        try:
            self.root.config(cursor="wait")
            self.root.update()
            
            result_df = self._search_engine.search(
                self._loader1._dataframe,
                self.selected_src_col,
                self._loader2._dataframe,
                self.selected_tgt_col,
                list(self.selected_output_cols)
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