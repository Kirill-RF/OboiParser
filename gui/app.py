"""Модуль графического интерфейса приложения."""
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, List

import pandas as pd

# Импорты из core модулей (предполагается, что структура проекта соблюдена)
# Если всё ещё в одном файле, эти импорты нужно убрать или закомментировать
from core.extractor import ArticleExtractor
from core.template_manager import TemplateManager
from core.data_loader import ExcelDataLoader
from core.search_engine import ArticleSearchEngine


class ArticleFinderGUI:
    """Главное окно приложения для поиска артикулов в Excel."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Поиск строк по артикулам")
        self.root.geometry("950x750")
        self.root.resizable(True, True)

        self._template_manager = TemplateManager()
        default_pattern = self._template_manager.get_current_pattern() or r'[A-Za-zА-Яа-я0-9\-_]{3,}'
        self._extractor = ArticleExtractor(default_pattern)
        self._search_engine = ArticleSearchEngine(self._extractor)
        
        self._loader1: Optional[ExcelDataLoader] = None
        self._loader2: Optional[ExcelDataLoader] = None

        self.selected_src_col: Optional[str] = None
        self.selected_tgt_col: Optional[str] = None
        self.selected_output_cols: List[str] = []
        # Колонки из Файла 1 для вывода
        self.selected_src_output_cols: List[str] = []  

        # Виджеты
        self.lbl_template: ttk.Label = None
        self.preview_tree1: ttk.Treeview = None
        self.preview_tree2: ttk.Treeview = None
        self.result_tree: ttk.Treeview = None
        
        # Метки статуса
        self.lbl_src_status: ttk.Label = None
        self.lbl_tgt_status: ttk.Label = None
        self.lbl_out_status: ttk.Label = None
        self.lbl_src_output_status: ttk.Label = None
        
        # Меню
        self.context_menu: tk.Menu = None # Главное меню шаблонов
        self.result_context_menu: tk.Menu = None # Меню для результатов поиска

        self.status_var: tk.StringVar = None
        self.status_bar: ttk.Label = None

        self._setup_ui()
        self._create_context_menus()

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
        ttk.Label(frm_template, text="Шаблон поиска: ").pack(side="left", padx=(0, 5))
        self.lbl_template = ttk.Label(frm_template, text="По умолчанию", foreground="blue")
        self.lbl_template.pack(side="left")
        ttk.Button(frm_template, text="Загрузить шаблон", command=self._load_template_via_dialog).pack(side="left", padx=(10, 0))
        ttk.Button(frm_template, text="Сбросить", command=self._reset_template).pack(side="left", padx=(5, 0))

        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W) 
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # --- Файл 1 ---
        frm1 = ttk.LabelFrame(self.root, text="1. Файл-Справочник (Где лежат артикулы)")
        frm1.pack(fill="x", padx=10, pady=5)
        
        btn_frame1 = ttk.Frame(frm1)
        btn_frame1.pack(fill="x", padx=5, pady=5)
        self.path1_var = tk.StringVar()
        ttk.Entry(btn_frame1, textvariable=self.path1_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame1, text="Выбрать", command=lambda: self._select_file(self.path1_var)).pack(side="left", padx=2)
        ttk.Button(btn_frame1, text="Загрузить", command=self._load_file1).pack(side="left", padx=2)

        # Создаём фрейм для горизонтального размещения статусов
        status_frame = ttk.Frame(frm1)
        status_frame.pack(fill="x", padx=10, pady=(0, 5))

        self.lbl_src_status = ttk.Label(status_frame, text="Выбрана колонка: Не выбрана (ЛКМ по заголовку)", foreground="grey")
        self.lbl_src_status.pack(side="left", padx=(0, 20))

        self.lbl_src_output_status = ttk.Label(status_frame, text="Колонки справочника: Не выбраны (ПКМ)", foreground="grey")
        self.lbl_src_output_status.pack(side="left")
        self.preview_tree1 = self._create_preview_widget(frm1, file_num=1)

        # --- Файл 2 ---
        frm2 = ttk.LabelFrame(self.root, text="2. Файл для поиска (Где ищем совпадения)")
        frm2.pack(fill="x", padx=10, pady=5)
        
        btn_frame2 = ttk.Frame(frm2)
        btn_frame2.pack(fill="x", padx=5, pady=5)
        self.path2_var = tk.StringVar()
        ttk.Entry(btn_frame2, textvariable=self.path2_var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(btn_frame2, text="Выбрать", command=lambda: self._select_file(self.path2_var)).pack(side="left", padx=2)
        ttk.Button(btn_frame2, text="Загрузить", command=self._load_file2).pack(side="left", padx=2)

        status_frame2 = ttk.Frame(frm2)
        status_frame2.pack(fill="x", padx=10, pady=(0, 5))
        self.lbl_tgt_status = ttk.Label(status_frame2, text="Колонка поиска: Не выбрана (ЛКМ)", foreground="grey")
        self.lbl_tgt_status.pack(side="left", padx=(0, 20))
        self.lbl_out_status = ttk.Label(status_frame2, text="Колонки вывода: Не выбраны (ПКМ)", foreground="grey")
        self.lbl_out_status.pack(side="left", padx=(0, 20))
        self.preview_tree2 = self._create_preview_widget(frm2, file_num=2)

        # --- Кнопки ---
        frm_btn = ttk.Frame(self.root)
        frm_btn.pack(pady=10)
        ttk.Button(frm_btn, text="Найти совпадения", command=self._find_matches).pack(side="left", padx=5)
        ttk.Button(frm_btn, text="Экспорт в Excel", command=self._export_results).pack(side="left", padx=5)
        ttk.Button(frm_btn, text="Очистить", command=self._clear_results).pack(side="left", padx=5)

        # --- Результаты ---
        frm_res = ttk.LabelFrame(self.root, text="Результаты поиска")
        frm_res.pack(fill="both", expand=True, padx=10, pady=5)
        tree_frame = ttk.Frame(frm_res)
        tree_frame.pack(fill="both", expand=True)
        
        self.result_tree = ttk.Treeview(tree_frame, show="headings")
        # ВАЖНО: Привязываем новое меню только к этому дереву
        self.result_tree.bind("<Button-3>", self._show_result_context_menu)
        
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.result_tree.xview)
        self.result_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        self.result_tree.pack(side="left", fill="both", expand=True)

    def _create_preview_widget(self, parent: ttk.Widget, file_num: int) -> ttk.Treeview:
        preview_frame = ttk.Frame(parent)
        preview_frame.pack(fill="both", expand=False, padx=5, pady=(0, 5))
        ttk.Label(preview_frame, text="Предпросмотр: ").pack(anchor="w")
        
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill="x", expand=True)
        
        tree = ttk.Treeview(tree_frame, columns=(), show="headings", height=2)
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=scroll_x.set)
        
        # События выбора колонок - ЛКМ и ПКМ мышью
        tree.bind('<Button-1>', lambda event: self._on_header_click(event, file_num, tree))
        tree.bind('<Button-3>', lambda event: self._on_header_rclick_extended(event, file_num, tree))
        
        tree.pack(side="left", fill="x", expand=True)
        scroll_x.pack(side="bottom", fill="x")
        return tree
    
    def _on_header_rclick_extended(self, event: tk.Event, file_num: int, tree: ttk.Treeview) -> None:
        """ПКМ: Добавление/Удаление колонки из списка вывода (для обоих файлов)."""
        
        region = tree.identify_region(event.x, event.y)
        if region != "heading": return

        col_id = tree.identify_column(event.x)
        col_index = int(col_id.lstrip("#")) - 1
        cols = tree["columns"]
        if col_index < 0 or col_index >= len(cols): return

        selected_col_name = cols[col_index]

        if file_num == 1:
            # Работа с колонками справочника
            # ✅ Разрешаем добавлять/удалять ЛЮБЫЕ колонки, включая колонку с артикулами
            if selected_col_name in self.selected_src_output_cols:
                self.selected_src_output_cols.remove(selected_col_name)
            else:
                self.selected_src_output_cols.append(selected_col_name)
            
            # Обновляем статус
            if self.selected_src_output_cols:
                names = [self._format_column_name(c, list(cols).index(c)) 
                        for c in self.selected_src_output_cols]
                self.lbl_src_output_status.config(
                    text=f"Колонки справочника: {', '.join(names)}", 
                    foreground="black"
                )
            else:
                self.lbl_src_output_status.config(
                    text="Колонки справочника: Не выбраны (ПКМ)", 
                    foreground="grey"
                )
            
            self._highlight_header_src_output(tree, cols)
            
        else:
            # Существующая логика для Файла 2
            if selected_col_name in self.selected_output_cols:
                self.selected_output_cols.remove(selected_col_name)
            else:
                self.selected_output_cols.append(selected_col_name)

            if self.selected_output_cols:
                names = []
                for c in self.selected_output_cols:
                    idx = list(cols).index(c)
                    names.append(self._format_column_name(c, idx))
                self.lbl_out_status.config(
                    text=f"Колонки вывода: {', '.join(names)}", 
                    foreground="black"
                )
            else:
                self.lbl_out_status.config(
                    text="Колонки вывода: Не выбраны (ПКМ)", 
                    foreground="grey"
                )

            self._highlight_header_output(tree, cols)
            
    def _highlight_header_src_output(self, tree: ttk.Treeview, cols: tuple) -> None:
        """Подсветка заголовков для Файла 1 (артикулы + дополнительные колонки)."""
        current_article_col = self.selected_src_col
        
        for col in cols:
            idx = list(cols).index(col)
            base_text = self._format_column_name(col, idx)
            
            is_output = col in self.selected_src_output_cols
            is_article = col == current_article_col
            
            if is_article and is_output:
                tree.heading(col, text=f"▶ ★ {base_text}")
            elif is_article:
                tree.heading(col, text=f"▶ {base_text}")
            elif is_output:
                tree.heading(col, text=f"★ {base_text}")
            else:
                tree.heading(col, text=base_text)
        
    def _create_context_menus(self) -> None:
        # 1. Главное контекстное меню (для пустых мест)
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Открыть каталог шаблонов", command=self._open_templates_dir)
        self.context_menu.add_command(label="Загрузить шаблон...", command=self._load_template_via_dialog)
        self.context_menu.add_command(label="Сбросить к стандартному", command=self._reset_template)
        self.root.bind("<Button-3>", self._show_context_menu)

        # 2. Меню для таблицы результатов (Копирование)
        self.result_context_menu = tk.Menu(self.root, tearoff=0)
        self.result_context_menu.add_command(label="Копировать выделенные строки", command=self._copy_selected_rows)
        self.result_context_menu.add_command(label="Копировать все строки", command=self._copy_all_rows)

    def _show_result_context_menu(self, event: tk.Event) -> None:
        """Показывает меню копирования для таблицы результатов."""
        try:
            self.result_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.result_context_menu.grab_release()

    def _show_context_menu(self, event: tk.Event) -> None:
        """Показывает главное меню (шаблоны), если клик не по таблице результатов."""
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        
        # Если клик был по таблице результатов - ничего не делаем (там своё меню)
        if widget == self.result_tree:
            return

        # Проверка на вложенность (на случай, если клик по скроллбару результатов)
        parent = widget
        while parent:
            if parent == self.result_tree:
                return
            parent = parent.master if hasattr(parent, 'master') else None

        # Также не показываем, если клик по деревьям предпросмотра (у них своя логика в других методах, 
        # но здесь мы просто блокируем всплывание главного меню)
        if widget == self.preview_tree1 or widget == self.preview_tree2:
            return

        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def _copy_to_clipboard(self, text: str) -> None:
        """Копирует текст в буфер обмена."""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update() # Чтобы буфер обновился даже после закрытия окна (опционально)
        self.status_var.set("Скопировано в буфер обмена!")

    def _copy_selected_rows(self) -> None:
        """Копирует выделенные пользователем строки."""
        selected_items = self.result_tree.selection()
        if not selected_items:
            self.status_var.set("Ничего не выделено.")
            return

        text = self._get_tree_data_as_text(self.result_tree, selected_items)
        self._copy_to_clipboard(text)
        self.status_var.set(f"Скопировано {len(selected_items)} строк.")

    def _copy_all_rows(self) -> None:
        """Копирует все строки из таблицы результатов."""
        all_items = self.result_tree.get_children()
        if not all_items:
            self.status_var.set("Таблица пуста.")
            return

        text = self._get_tree_data_as_text(self.result_tree, all_items)
        self._copy_to_clipboard(text)
        self.status_var.set(f"Скопировано {len(all_items)} строк.")

    def _open_templates_dir(self) -> None:
        templates_dir = self._template_manager.get_templates_dir()
        if os.name == 'nt':
            os.startfile(templates_dir)
        elif os.name == 'posix':
            import subprocess
            subprocess.call(['xdg-open', templates_dir])

    def _load_template_via_dialog(self) -> None:
        templates_dir = self._template_manager.get_templates_dir()
        filepath = filedialog.askopenfilename(
            initialdir=templates_dir, 
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if filepath:
            self._load_template(filepath)

    def _load_template(self, template_path: str) -> None:
        pattern = self._template_manager.load_template(template_path)
        if pattern:
            try:
                self._extractor = ArticleExtractor(pattern)
                self._search_engine = ArticleSearchEngine(self._extractor)
                self.lbl_template.config(
                    text=f"{self._template_manager.get_current_template_name()}", 
                    foreground="green"
                )
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
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel файлы", "*.xlsx *.xls"), ("Все файлы", "*.*")]
        )
        if filepath:
            string_var.set(filepath)

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
                self.selected_src_output_cols = []  # Сброс
                self.lbl_src_status.config(text="Выбрана колонка: Не выбрана (ЛКМ по заголовку)", 
                                           foreground="grey")
                if hasattr(self, 'lbl_src_output_status'):
                    self.lbl_src_output_status.config(text="Колонки справочника: Не выбраны (ПКМ)",
                                                      foreground="grey")
                self._update_preview(self.preview_tree1, loader.get_first_row(), 1)
            else:
                self._loader2 = loader
                self.selected_tgt_col = None
                self.selected_output_cols = []
                self.lbl_tgt_status.config(text="Колонка поиска: Не выбрана (ЛКМ)",
                                           foreground="grey")
                self.lbl_out_status.config(text="Колонки вывода: Не выбраны (ПКМ)", 
                                           foreground="grey")
                self._update_preview(self.preview_tree2, loader.get_first_row(), 2)

            # ✅ Показываем сообщение в области результатов
            self._show_message_in_results(f"Файл {file_num} загружен: {os.path.basename(path)}")
            # ✅ И дублируем в строку состояния для надёжности
            self.status_var.set(f"Файл {file_num} загружен: {os.path.basename(path)}")
            
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    def _on_header_click(self, event: tk.Event, file_num: int, tree: ttk.Treeview) -> None:
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
            self.lbl_src_status.config(text=f"Выбрана колонка: {display_name}", foreground="black")
            self._highlight_header_single(tree, cols, selected_col_name)
        else:
            self.selected_tgt_col = selected_col_name
            display_name = self._format_column_name(selected_col_name, col_index)
            self.lbl_tgt_status.config(text=f"Колонка поиска: {display_name}", foreground="black")
            self._highlight_header_search(tree, cols, selected_col_name)

    def _on_header_rclick(self, event: tk.Event, file_num: int, tree: ttk.Treeview) -> None:
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
            self.selected_output_cols.append(selected_col_name)

        if self.selected_output_cols:
            names = []
            for c in self.selected_output_cols:
                idx = list(cols).index(c)
                names.append(self._format_column_name(c, idx))
            self.lbl_out_status.config(text=f"Колонки вывода: {', '.join(names)}", foreground="black")
        else:
            self.lbl_out_status.config(text="Колонки вывода: Не выбраны (ПКМ)", foreground="grey")

        self._highlight_header_output(tree, cols)

    def _highlight_header_single(self, tree: ttk.Treeview, cols: tuple, selected_col: str) -> None:
        for col in cols:
            idx = list(cols).index(col)
            tree.heading(col, text=self._format_column_name(col, idx))
        idx = list(cols).index(selected_col)
        tree.heading(selected_col, text=f"▶ {self._format_column_name(selected_col, idx)}")

    def _highlight_header_search(self, tree: ttk.Treeview, cols: tuple, selected_col: str) -> None:
        for col in cols:
            idx = list(cols).index(col)
            base_text = self._format_column_name(col, idx)
            if col in self.selected_output_cols:
                tree.heading(col, text=f"★ {base_text}")
            else:
                tree.heading(col, text=base_text)
        idx = list(cols).index(selected_col)
        base_text = self._format_column_name(selected_col, idx)
        tree.heading(selected_col, text=f"▶ {base_text}")

    def _highlight_header_output(self, tree: ttk.Treeview, cols: tuple) -> None:
        current_search = self.selected_tgt_col
        for col in cols:
            idx = list(cols).index(col)
            base_text = self._format_column_name(col, idx)
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
             # ✅ Показываем статус в области результатов
            # self._show_message_in_results("🔍 Выполняется поиск...", duration_ms=10000)
            
            # ✅ Очищаем предыдущие результаты перед новым поиском
            self._clear_results()
            self.status_var.set("Выполняется поиск...")
            
            result_df = self._search_engine.search(
                self._loader1.get_dataframe(),  # ✅ Используем геттер вместо _dataframe
                self.selected_src_col,
                self._loader2.get_dataframe(),
                self.selected_tgt_col,
                self.selected_output_cols,
                self.selected_src_output_cols
            )
            
            self._display_results(result_df)
            
            if result_df.empty:
                self.status_var.set("Совпадений не найдено.")
            else:
                self.status_var.set(f"Найдено строк: {len(result_df)}")
                                
        except Exception as e:
            messagebox.showerror("Ошибка поиска", str(e))
        finally:
            self.root.config(cursor="")

    def _clear_results(self) -> None:
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.result_tree["columns"] = ()
        self.status_var.set('')

    def _export_results(self) -> None:
        """Экспорт результатов в Excel."""
        items = self.result_tree.get_children()
        if not items:
            self.status_var.set("Нет данных для экспорта.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
            title="Сохранить результаты поиска"
        )
        if file_path:
            try:
                cols = self.result_tree["columns"]
                # Форматируем имена колонок
                formatted_cols = [self._format_column_name(col, idx) for idx, col in enumerate(cols)]
                
                data = [self.result_tree.item(item, "values") for item in items]
                df = pd.DataFrame(data, columns=formatted_cols)
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Успех", f"Результаты сохранены в:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Ошибка экспорта", f"Не удалось сохранить файл:\n{e}")
                
    def _get_tree_data_as_text(self, tree: ttk.Treeview, items: tuple) -> str:
        """Преобразует строки Treeview в текстовый формат (TSV)."""
        if not items:
            return ""
        
        # Заголовки колонок с форматированием
        cols = tree["columns"]
        formatted_headers = [self._format_column_name(col, idx) for idx, col in enumerate(cols)]
        header = "\t".join(formatted_headers)
        
        # Данные
        rows_text = []
        for item_id in items:
            values = tree.item(item_id, "values")
            row_str = "\t".join(str(val) for val in values)
            rows_text.append(row_str)
            
        return header + "\n" + "\n".join(rows_text)

    def _display_results(self, df: pd.DataFrame) -> None:
        self._clear_results()
        if df.empty: return

        cols = df.columns.astype(str).tolist()
        self.result_tree["columns"] = cols
        for i, col in enumerate(cols):
            display_name = self._format_column_name(col, i)
            self.result_tree.heading(col, text=display_name)
            max_len = df[col].astype(str).map(len).max()
            width = max(50, min(300, max_len * 7 + 10))
            self.result_tree.column(col, width=width, anchor="w")

        for row in df.values.tolist():
            self.result_tree.insert("", "end", values=[str(val) for val in row])
            
    def _show_message_in_results(self, message: str, duration_ms: int = 3000) -> None:
        """
        Отображает временное информационное сообщение в таблице результатов.
        
        Parameters
        ----------
        message : str
            Текст сообщения для отображения.
        duration_ms : int, optional
            Время показа сообщения в миллисекундах (по умолчанию 3000 мс = 3 сек).
        """
        # Очищаем таблицу
        self._clear_results()
        
        # Настраиваем одну колонку для сообщения
        self.result_tree["columns"] = ("message",)
        self.result_tree.heading("message", text="ℹ️ Информация")
        self.result_tree.column("message", width=400, anchor="center")
        
        # Добавляем строку с сообщением
        self.result_tree.insert("", "end", values=(message,))
        
        # Автоматически очищаем сообщение через указанное время
        self.root.after(duration_ms, self._clear_results)