import os
import tempfile
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from typing import List


class ChartBuilder:
    """Генерирует изображения диаграмм и возвращает пути к временным файлам."""

    def __init__(self):
        self._temp_files = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
        return False

    def _save_and_close(self, fig) -> str:
        tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        fig.savefig(tmp.name, dpi=150, bbox_inches='tight')
        plt.close(fig)
        self._temp_files.append(tmp.name)
        return tmp.name

    def cleanup(self):
        for path in self._temp_files:
            try:
                os.unlink(path)
            except OSError:
                pass
        self._temp_files.clear()

    def create_pie_chart(self, labels: List[str], values: List[int],
                         title: str = '') -> str:
        fig, ax = plt.subplots(figsize=(6, 4))
        wedges, texts, autotexts = ax.pie(
            values,
            labels=labels,
            autopct='%1.1f%%',
            startangle=90,
            textprops={'fontsize': 10}
        )
        ax.set_title(title, fontsize=12)
        return self._save_and_close(fig)

    def create_line_chart(self,
                          x_labels: List[str],
                          y_values: List[int],
                          xlabel: str = 'Месяц',
                          ylabel: str = 'Количество документов',
                          title: str = '') -> str:
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.plot(x_labels, y_values, marker='o', linestyle='-', color='#1f77b4')
        ax.set_xlabel(xlabel)
        ax.set_ylabel(ylabel)
        ax.set_title(title)
        ax.grid(True, linestyle='--', alpha=0.7)
        plt.xticks(rotation=45)
        return self._save_and_close(fig)

    def create_horizontal_bar_chart(self,
                                    categories: List[str],
                                    values: List[int],
                                    title: str = '',
                                    xlabel: str = 'Количество ошибок категории 1') -> str:
        fig, ax = plt.subplots(figsize=(8, 5))
        y_pos = range(len(categories))
        ax.barh(y_pos, values, align='center', color='#d62728')
        ax.set_yticks(y_pos)
        ax.set_yticklabels(categories)
        ax.invert_yaxis()
        ax.set_xlabel(xlabel)
        ax.set_title(title)
        return self._save_and_close(fig)

    def create_monthly_trend_chart(self,
                                   months: List[str],
                                   errors_cat1: List[int],
                                   errors_cat2: List[int],
                                   title: str = 'Динамика ошибок по месяцам') -> str:
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.plot(months, errors_cat1, marker='o', label='Категория 1', color='red')
        ax.plot(months, errors_cat2, marker='s', label='Категория 2', color='blue')
        ax.set_xlabel('Месяц')
        ax.set_ylabel('Количество ошибок')
        ax.set_title(title)
        ax.grid(True, alpha=0.3)
        ax.legend()
        plt.xticks(rotation=45)
        return self._save_and_close(fig)

    def create_vertical_bar_chart(self,
                                 categories: List[str],
                                 values: List[int],
                                 title: str = '',
                                 xlabel: str = '',
                                 ylabel: str = '') -> str:
        """Вертикальная столбчатая диаграмма."""
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(categories, values, color='#4472C4')
        ax.set_xlabel(xlabel)
        ax.set_ylabel(ylabel)
        ax.set_title(title)
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        return self._save_and_close(fig)