"""Модуль поиска и фильтрации строк по артикулам."""
from typing import List, Optional
import pandas as pd
from .extractor import ArticleExtractor


class ArticleSearchEngine:
    """Движок поиска строк по списку артикулов."""

    def __init__(self, extractor: ArticleExtractor) -> None:
        self._extractor = extractor

    def search(self, source_df: pd.DataFrame, source_col: str, 
               target_df: pd.DataFrame, target_col: str, 
               output_cols: List[str],
               source_output_cols: Optional[List[str]] = None) -> pd.DataFrame:
        """Ищет совпадения артикулов и возвращает объединённую таблицу.

        Parameters
        ----------
        source_df : pd.DataFrame
            DataFrame-справочник.
        source_col : str
            Имя колонки в source_df с артикулами.
        target_df : pd.DataFrame
            DataFrame для поиска.
        target_col : str
            Имя колонки в target_df для сравнения.
        output_cols : List[str]
            Список колонок из target_df для вывода.
        source_output_cols : List[str], optional
            Список дополнительных колонок из source_df для вывода.

        Returns
        -------
        pd.DataFrame
            Объединённый DataFrame с результатами.
        """
        if source_output_cols is None:
            source_output_cols = []

        # Собираем эталонные артикулы
        source_articles = set()
        for val in source_df[source_col]:
            source_articles.update(self._extractor.extract(str(val)))

        source_articles.discard('')
        source_articles.discard('nan')
        if not source_articles:
            raise ValueError("В колонке справочника не найдено артикулов.")

        # Фильтруем второй файл
        mask = target_df[target_col].apply(
            lambda x: bool(set(self._extractor.extract(str(x))).intersection(source_articles))
        )
        filtered_df = target_df.loc[mask].copy()

        if filtered_df.empty:
            return pd.DataFrame()

        # Если есть дополнительные колонки из справочника
        if source_output_cols:
            # Создаём маппинг артикул -> данные из справочника
            article_map = {}
            for _, row in source_df.iterrows():
                articles = self._extractor.extract(str(row[source_col]))
                for art in articles:
                    if art and art not in article_map:
                        article_map[art] = {
                            col: row[col] for col in source_output_cols 
                            if col in source_df.columns
                        }

            # Добавляем колонки из справочника в результаты
            for src_col in source_output_cols:
                if src_col in source_df.columns:
                    new_name = f"Справочник.{src_col}"
                    filtered_df[new_name] = filtered_df[target_col].apply(
                        lambda x: next(
                            (article_map[art].get(src_col, "") 
                             for art in self._extractor.extract(str(x)) 
                             if art in article_map),
                            ""
                        )
                    )
                    # Вставляем новую колонку после первой колонки вывода
                    cols_list = filtered_df.columns.tolist()
                    if new_name in cols_list:
                        cols_list.remove(new_name)
                        # Находим позицию первой колонки и вставляем после неё
                        first_col_pos = 0
                        for i, col in enumerate(cols_list):
                            if col in output_cols:
                                first_col_pos = i
                                break
                        cols_list.insert(first_col_pos + 1, new_name)
                        filtered_df = filtered_df[cols_list]

        # Выбираем нужные колонки в правильном порядке
        final_cols = []
        
        # Сначала добавляем колонки из справочника (если есть)
        for src_col in source_output_cols:
            new_name = f"Справочник.{src_col}"
            if new_name in filtered_df.columns:
                final_cols.append(new_name)
        
        # Затем добавляем выбранные колонки из целевого файла
        for col in output_cols:
            if col in filtered_df.columns:
                final_cols.append(col)
        
        # Если ничего не выбрано, берём все колонки
        if not final_cols:
            final_cols = filtered_df.columns.tolist()

        return filtered_df[final_cols]