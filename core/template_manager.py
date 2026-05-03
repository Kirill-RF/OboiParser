"""Модуль управления шаблонами регулярных выражений."""
import os
from typing import Optional
import tkinter.messagebox as messagebox


class TemplateManager:
    """Менеджер жизненного цикла шаблонов поиска.

    Отвечает за создание директории, чтение файлов конфигурации,
    хранение состояния текущего шаблона и сброс к значениям по умолчанию.

    Attributes
    ----------
    _templates_dir : str
        Абсолютный или относительный путь к папке с шаблонами.
    _current_pattern : str or None
        Строка текущего регулярного выражения.
    _current_template_name : str or None
        Имя файла последнего загруженного шаблона.
    """

    def __init__(self, templates_dir: str = "templates") -> None:
        self._templates_dir = templates_dir
        self._current_pattern: Optional[str] = None
        self._current_template_name: Optional[str] = None
        self._ensure_templates_dir()

    def _ensure_templates_dir(self) -> None:
        """Создаёт директорию и файл-пример, если они отсутствуют."""
        if not os.path.exists(self._templates_dir):
            os.makedirs(self._templates_dir)
        
        example_path = os.path.join(self._templates_dir, "example_patterns.txt")
        if not os.path.exists(example_path):
            with open(example_path, 'w', encoding='utf-8') as f:
                f.write("# Шаблон для артикулов обоев\n")
                f.write(r"\b(?:[A-Z]{0,3}\d{3,6}(?:[-–]\d{1,3})?|\d{4,6}(?:[-–][A-Z0-9]{1,3})?)\b")

    def get_templates_dir(self) -> str:
        """Возвращает путь к директории шаблонов.

        Returns
        -------
        str
            Путь к папке `templates`.
        """
        return self._templates_dir

    def load_template(self, template_path: str) -> Optional[str]:
        """Загружает шаблон из текстового файла.

        Parameters
        ----------
        template_path : str
            Полный путь к `.txt` файлу.

        Returns
        -------
        str or None
            Строка регулярного выражения или `None` при ошибке.
        """
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
            messagebox.showerror("Ошибка", str(e))
            return None

    def get_current_pattern(self) -> Optional[str]:
        return self._current_pattern

    def get_current_template_name(self) -> Optional[str]:
        return self._current_template_name

    def reset_to_default(self) -> None:
        """Сбрасывает шаблон к настройкам по умолчанию."""
        self._current_pattern = r'[A-Za-zА-Яа-я0-9\-_]{3,}'
        self._current_template_name = "По умолчанию"