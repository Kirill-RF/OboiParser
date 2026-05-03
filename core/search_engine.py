"""Модуль поиска и фильтрации строк по артикулам."""
from typing import List
import pandas as pd
from .extractor import ArticleExtractor


class ArticleSearchEngine:
    """Движок поиска строк по списку артикулов.

    Использует паттерн Dependency Injection для получения экстрактора,
    что позволяет легко менять алгоритм парсинга без изменения логики поиска (DIP).

    Attributes
    ----------
    _extractor : ArticleExtractor
        Экземпляр класса для извлечения артикулов из текста.
    """

    def __init__(self, extractor: ArticleExtractor) -> None:
        self._extractor = extractor

    def search(self, source_df: pd.DataFrame, source_col: str, 
               target_df: pd.DataFrame, target_col: str, output_cols: List[str]) -> pd.DataFrame:
        """Ищет совпадения артикулов и возвращает отфильтрованную таблицу.

        Parameters
        ----------
        source_df : pd.DataFrame
            DataFrame-справочник, содержащий эталонные артикулы.
        source_col : str
            Имя колонки в `source_df` для извлечения эталонов.
        target_df : pd.DataFrame
            DataFrame, в котором выполняется поиск.
        target_col : str
            Имя колонки в `target_df` для сравнения.
        output_cols : List[str]
            Список колонок, которые нужно оставить в результате (в порядке выбора).

        Returns
        -------
        pd.DataFrame
            Отфильтрованный DataFrame. Пустой DataFrame, если совпадений нет.

        Raises
        ------
        ValueError
            Если в колонке-источнике не найдено ни одного валидного артикула.
        """
        source_articles = set()
        for val in source_df[source_col]:
            source_articles.update(self._extractor.extract(str(val)))
            
        source_articles.discard('')
        source_articles.discard('nan')
        if not source_articles:
            raise ValueError("В колонке справочника не найдено артикулов.")

        mask = target_df[target_col].apply(
            lambda x: bool(set(self._extractor.extract(str(x))).intersection(source_articles))
        )
        filtered_df = target_df.loc[mask].copy()

        if filtered_df.empty:
            return pd.DataFrame()

        valid_cols = [col for col in output_cols if col in filtered_df.columns]
        if not valid_cols:
            valid_cols = filtered_df.columns.tolist()
            
        return filtered_df[valid_cols]