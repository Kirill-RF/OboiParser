"""Модуль загрузки и предварительной обработки табличных данных."""
import os
from typing import List, Optional
import pandas as pd


class ExcelDataLoader:
    """Класс-загрузчик Excel файлов.

    Инкапсулирует логику выбора движка `openpyxl`/`xlrd` и базовой очистки данных.
    Реализует принцип открытости/закрытости (OCP): при необходимости поддержки
    новых форматов достаточно добавить ветвь в `load()`, не меняя интерфейс.

    Attributes
    ----------
    _file_path : str
        Путь к файлу.
    _dataframe : pd.DataFrame or None
        Кэш загруженных данных.
    """

    def __init__(self, file_path: str) -> None:
        self._file_path = file_path
        self._dataframe: Optional[pd.DataFrame] = None

    def load(self) -> pd.DataFrame:
        """Загружает Excel файл и возвращает очищенный DataFrame.

        Returns
        -------
        pd.DataFrame
            Таблица с заменёнными `NaN` на пустые строки.

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
        engine_map = {".xlsx": "openpyxl", ".xls": "xlrd"}
        
        if ext not in engine_map:
            raise ValueError("Поддерживаются только форматы .xls и .xlsx")

        self._dataframe = pd.read_excel(self._file_path, engine=engine_map[ext])
        self._dataframe = self._dataframe.fillna('')
        return self._dataframe

    def get_columns(self) -> List[str]:
        """Возвращает список названий колонок.

        Returns
        -------
        List[str]
            Названия колонок в виде строк.
        """
        if self._dataframe is None:
            return []
        return self._dataframe.columns.astype(str).tolist()

    def get_first_row(self) -> Optional[pd.Series]:
        """Возвращает первую строку для предпросмотра.

        Returns
        -------
        pd.Series or None
            Первая строка данных или `None`, если таблица пуста.
        """
        if self._dataframe is None or self._dataframe.empty:
            return None
        return self._dataframe.iloc[0]

    def get_dataframe(self) -> pd.DataFrame:
        """Возвращает загруженный DataFrame (без прямого доступа к приватному атрибуту).

        Returns
        -------
        pd.DataFrame

        Raises
        ------
        RuntimeError
            Если данные ещё не загружены.
        """
        if self._dataframe is None:
            raise RuntimeError("Данные ещё не загружены. Сначала вызовите метод load().")
        return self._dataframe