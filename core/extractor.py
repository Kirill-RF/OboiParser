"""Модуль для извлечения структурированных данных (артикулов) из текста."""
import re
from typing import List, Optional


class ArticleExtractor:
    """Класс для извлечения артикулов из текстовых строк с помощью регулярных выражений.

    Реализует принцип единственной ответственности (SRP): занимается только
    парсингом текста согласно заданному шаблону.

    Attributes
    ----------
    _pattern : re.Pattern
        Скомпилированное регулярное выражение для поиска.
    """

    def __init__(self, pattern: str = r'[A-Za-zА-Яа-я0-9\-_]{3,}') -> None:
        """Инициализация экстрактора.

        Parameters
        ----------
        pattern : str, optional
            Строка регулярного выражения. По умолчанию: `r'[A-Za-zА-Яа-я0-9\-_]{3,}'`
        """
        self._pattern = re.compile(pattern)

    def extract(self, text: Optional[str]) -> List[str]:
        """Извлекает уникальные артикулы из текстовой строки.

        Parameters
        ----------
        text : str or None
            Исходная строка для анализа.

        Returns
        -------
        List[str]
            Список уникальных найденных артикулов. Возвращает пустой список,
            если `text` не является строкой или пуст.
        """
        if not isinstance(text, str) or not text.strip():
            return []
        return list(set(self._pattern.findall(text)))