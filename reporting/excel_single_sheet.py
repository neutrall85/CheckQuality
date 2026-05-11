import os
from datetime import datetime
from typing import List, Dict, Tuple
import xlsxwriter
from core.models.statistics import Statistics
from core.models.config import ConfigModel
from .base_report import BaseReport
from .styles import *

RU_MONTHS = [
    '', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
]

SECTION_TITLE_STYLE = {
    'bold': True,
    'font_size': 12,
    'font_color': 'white',
    'bg_color': '#4472C4',
    'valign': 'vcenter',
    'text_wrap': True
}

class ExcelSingleSheetReport(BaseReport):
    """Генератор сводного отчёта Excel – 8 диаграмм (добавлена гистограмма по группам файлов)."""

    def __init__(self):
        self.workbook = None

    def generate(self, statistics: Statistics, config: ConfigModel, output_path: str) -> str:
        self.workbook = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False})
        ws = self.workbook.add_worksheet('Сводный отчёт')

        ws.set_landscape()
        ws.set_paper(9)
        ws.set_footer(f'&CДата формирования: {datetime.now():%d.%m.%Y %H:%M}')

        header_fmt = self.workbook.add_format(FMT_HEADER)
        cell_wrap = self.workbook.add_format(FMT_CELL_WRAP)
        bold_fmt = self.workbook.add_format(FMT_BOLD)
        pct_fmt = self.workbook.add_format({**FMT_CELL_WRAP, **FMT_PERCENT})
        num_fmt = self.workbook.add_format({**FMT_CELL_WRAP, 'num_format': '0.00'})
        section_title_fmt = self.workbook.add_format(SECTION_TITLE_STYLE)
        total_fmt = self.workbook.add_format({**FMT_CELL_WRAP, 'bold': True, 'num_format': '0.00'})
        total_fmt_int = self.workbook.add_format({**FMT_CELL_WRAP, 'bold': True})

        chart_ws = self.workbook.add_worksheet('_chart_data')
        chart_ws.hide()
        num_months, num_devs, num_groups = self._write_chart_data(chart_ws, statistics)

        row = 0
        row = self._zone1(ws, statistics, config, bold_fmt, cell_wrap, section_title_fmt)
        row = self._zone2(ws, statistics, row, header_fmt, cell_wrap, pct_fmt, num_fmt,
                          section_title_fmt, total_fmt, total_fmt_int)
        row = self._zone3(ws, statistics, row, header_fmt, cell_wrap, num_fmt,
                          section_title_fmt, total_fmt, total_fmt_int)
        row = self._zone6(ws, statistics, row, header_fmt, cell_wrap, num_fmt,
                          section_title_fmt, total_fmt, total_fmt_int)
        self._zone7(ws, statistics, row, bold_fmt, cell_wrap, num_fmt, pct_fmt,
                     section_title_fmt)

        ws.set_column('A:A', 35)
        for col in range(1, 12):
            ws.set_column(col, col, 12)

        self._insert_charts(ws, statistics, num_months, num_devs, num_groups)

        self.workbook.close()
        return output_path

    # -----------------------------------------------------------------
    # Скрытый лист с данными для ВСЕХ диаграмм
    # -----------------------------------------------------------------
    def _write_chart_data(self, ws, stats):
        # 1. Круговая "Распределение ошибок" (взаимоисключающие категории)
        ws.write('A1', 'Категория')
        ws.write('B1', 'Количество документов')
        docs_cat1 = stats.total_docs_with_only_cat1
        docs_cat2 = stats.total_docs_with_only_cat2
        docs_both = stats.total_docs_with_both
        docs_no_errors = stats.total_docs - stats.docs_with_errors
        ws.write('A2', 'Только кат.1')
        ws.write('B2', docs_cat1)
        ws.write('A3', 'Только кат.2')
        ws.write('B3', docs_cat2)
        ws.write('A4', 'Обе категории')
        ws.write('B4', docs_both)
        ws.write('A5', 'Без ошибок')
        ws.write('B5', docs_no_errors)

        # 2. Круговая "Доля документов с ошибками"
        ws.write('A7', 'Категория')
        ws.write('B7', 'Количество документов')
        ws.write('A8', 'С ошибками:')
        ws.write('B8', stats.docs_with_errors)
        ws.write('A9', 'Без ошибок:')
        ws.write('B9', docs_no_errors)

        # 3. Гистограмма "Количество документов по месяцам"
        ws.write('D1', 'Месяц')
        ws.write('E1', 'Количество документов')
        months = sorted(stats.by_month.keys())
        for i, m in enumerate(months, start=2):
            year, month_num = map(int, m.split('-'))
            month_name = f"{RU_MONTHS[month_num]} {year}"
            data = stats.by_month[m]
            ws.write(i-1, 3, month_name)
            ws.write(i-1, 4, data['count'])

        # 4. Гистограмма "Ошибки по месяцам"
        ws.write('G1', 'Месяц')
        ws.write('H1', 'Ошибки кат.1')
        ws.write('I1', 'Ошибки кат.2')
        for i, m in enumerate(months, start=2):
            year, month_num = map(int, m.split('-'))
            month_name = f"{RU_MONTHS[month_num]} {year}"
            data = stats.by_month[m]
            ws.write(i-1, 6, month_name)
            ws.write(i-1, 7, data['errors1'])
            ws.write(i-1, 8, data['errors2'])

        # 5-7. Гистограммы по разработчикам
        ws.write(0, 10, 'Разработчик')
        ws.write(0, 11, 'Кол-во документов')
        ws.write(0, 12, 'Ошибки кат.1')
        ws.write(0, 13, 'Ошибки кат.2')
        devs = sorted(stats.by_developer.items(), key=lambda x: x[1]['count'], reverse=True)
        for i, (dev, d) in enumerate(devs, start=1):
            ws.write(i, 10, dev)
            ws.write(i, 11, d['count'])
            ws.write(i, 12, d['errors1'])
            ws.write(i, 13, d['errors2'])

        # 8. Данные для гистограммы по группам файлов
        ws.write('P1', 'Группа файлов')
        ws.write('Q1', 'Количество документов')
        groups = sorted(stats.docs_by_file_prefix.items(), key=lambda x: x[1], reverse=True)
        for i, (group, count) in enumerate(groups, start=2):
            ws.write(i-1, 15, group)   # колонка P (индекс 15)
            ws.write(i-1, 16, count)   # колонка Q (индекс 16)

        return len(months), len(devs), len(groups)

    # -----------------------------------------------------------------
    # Зона 1 – Заголовок и общая информация
    # -----------------------------------------------------------------
    def _zone1(self, ws, stats, config, bold, cell, section_title_fmt):
        title_format = self.workbook.add_format({
            'bold': True,
            'font_size': 16,
            'font_name': 'Calibri',
            'font_color': 'black'
        })
        ws.write(0, 0, 'СТАТИСТИКА ЖУРНАЛА ОШИБОК', title_format)
        ws.write(1, 0, f'Дата формирования: {datetime.now():%d.%m.%Y %H:%M}')

        if config.period_start or config.period_end:
            start_str = config.period_start.strftime('%d.%m.%Y') if config.period_start else '...'
            end_str = config.period_end.strftime('%d.%m.%Y') if config.period_end else '...'
            p = f"с {start_str} по {end_str}"
        else:
            if stats.actual_start_date and stats.actual_end_date:
                p = f"с {stats.actual_start_date:%d.%m.%Y} по {stats.actual_end_date:%d.%m.%Y}"
            else:
                p = "За весь период"
        ws.write(2, 0, f'Период данных: {p}')
        ws.write(3, 0, f'Всего записей: {stats.total_docs}')
        row = 4
        if config.selected_developers:
            ws.write(row, 0, f'Отчёт по разработчикам: {", ".join(config.selected_developers)}')
            row += 1
        return row

    # -----------------------------------------------------------------
    # Зона 2 – Статистика по типам документов
    # -----------------------------------------------------------------
    def _zone2(self, ws, stats, start_row, hdr, cell, pct, num, section_title_fmt, total_fmt, total_fmt_int):
        current = start_row + 1
        ws.write(current, 0, 'СТАТИСТИКА ПО ТИПАМ ДОКУМЕНТОВ', section_title_fmt)
        current += 2
        headers = ['Тип документа','Количество','Всего ошибок кат.1','Всего ошибок кат.2',
                   'Средн. кол-во ошибок кат.1 на док','Средн. кол-во ошибок кат.2 на док',
                   'Всего форматов А4','Средн. кол-во форматов А4',
                   'Средн. кол-во ошибок кат.1 на А4','Средн. кол-во ошибок кат.2 на А4']
        for c, h in enumerate(headers):
            ws.write(current, c, h, hdr)
        current += 1

        total = stats.total_docs
        if total == 0:
            return current

        sorted_items = sorted(stats.by_type.items(), key=lambda x: x[1]['count'], reverse=True)
        for dtype, d in sorted_items:
            cnt = d['count']; e1 = d['errors1']; e2 = d['errors2']; a4 = d['a4']
            avg1 = e1 / cnt if cnt else 0
            avg2 = e2 / cnt if cnt else 0
            avg_a4 = a4 / cnt if cnt else 0
            a1_per_a4 = e1 / a4 if a4 else 0
            a2_per_a4 = e2 / a4 if a4 else 0
            ws.write(current, 0, dtype, cell)
            ws.write(current, 1, cnt, cell)
            ws.write(current, 2, e1, cell)
            ws.write(current, 3, e2, cell)
            ws.write(current, 4, avg1, num)
            ws.write(current, 5, avg2, num)
            ws.write(current, 6, a4, cell)
            ws.write(current, 7, avg_a4, num)
            ws.write(current, 8, a1_per_a4, num)
            ws.write(current, 9, a2_per_a4, num)
            current += 1

        # Итог
        ws.write(current, 0, 'ИТОГО ПО ВСЕМ ДОКУМЕНТАМ:', total_fmt_int)
        ws.write(current, 1, total, total_fmt_int)
        ws.write(current, 2, stats.total_errors_cat1, total_fmt_int)
        ws.write(current, 3, stats.total_errors_cat2, total_fmt_int)
        avg1 = stats.total_errors_cat1 / total if total else 0
        avg2 = stats.total_errors_cat2 / total if total else 0
        ws.write(current, 4, avg1, total_fmt)
        ws.write(current, 5, avg2, total_fmt)
        ws.write(current, 6, stats.total_a4, total_fmt_int)
        avg_a4 = stats.total_a4 / total if total else 0
        ws.write(current, 7, avg_a4, total_fmt)
        a1a4 = stats.total_errors_cat1 / stats.total_a4 if stats.total_a4 else 0
        a2a4 = stats.total_errors_cat2 / stats.total_a4 if stats.total_a4 else 0
        ws.write(current, 8, a1a4, total_fmt)
        ws.write(current, 9, a2a4, total_fmt)
        return current + 1

    # -----------------------------------------------------------------
    # Зона 3 – Рейтинг разработчиков + доп. информация + качество
    # -----------------------------------------------------------------
    def _zone3(self, ws, stats, start_row, hdr, cell, num, section_title_fmt, total_fmt, total_fmt_int):
        current = start_row + 1
        ws.write(current, 0, 'СТАТИСТИКА ПО РАЗРАБОТЧИКАМ (ВСЕ АВТОРЫ)', section_title_fmt)
        current += 2
        current = self._write_developer_rating_table(ws, stats, current, hdr, cell, num,
                                                     total_fmt, total_fmt_int)
        current += 1
        current = self._write_additional_info(ws, stats, current)
        current += 1
        current = self._write_quality_distribution(ws, stats, current)
        return current

    def _write_developer_rating_table(self, ws, stats, start_row, hdr, cell, num,
                                      total_fmt, total_fmt_int):
        """Таблица рейтинга разработчиков."""
        headers = ['Разработчик','Количество документов','Количество форматов А4',
                   'Всего ошибок кат.1','Всего ошибок кат.2',
                   'Средн. кол-во ошибок кат.1 на док','Средн. кол-во ошибок кат.2 на док',
                   'Средн. кол-во ошибок кат.1 на А4','Средн. кол-во ошибок кат.2 на А4',
                   'Коэф. ошибок','Рейтинг']
        for c, h in enumerate(headers):
            ws.write(start_row, c, h, hdr)
        current = start_row + 1

        devs = []
        for dev, d in stats.by_developer.items():
            cnt = d['count']
            if cnt == 0: continue
            a4 = d['a4']; e1 = d['errors1']; e2 = d['errors2']
            avg1 = e1 / cnt if cnt else 0
            avg2 = e2 / cnt if cnt else 0
            a1a4 = e1 / a4 if a4 else 0
            a2a4 = e2 / a4 if a4 else 0
            coef = (e1 + 0.2 * e2) / cnt if cnt else 0
            devs.append((dev, cnt, a4, e1, e2, avg1, avg2, a1a4, a2a4, coef))
        devs.sort(key=lambda x: x[5], reverse=True)

        for rank, it in enumerate(devs, start=1):
            avg1_val = it[5]
            if avg1_val >= THRESHOLD_POOR:
                color = COLOR_POOR
            elif avg1_val >= THRESHOLD_GOOD:
                color = COLOR_AVERAGE
            elif avg1_val >= THRESHOLD_EXCELLENT:
                color = COLOR_GOOD
            else:
                color = COLOR_EXCELLENT
            fmt_color = self.workbook.add_format({
                **FMT_CELL_WRAP, 'bg_color': color, 'border': 1, 'num_format': '0.00'
            })
            ws.write(current, 0, it[0], cell)
            ws.write(current, 1, it[1], cell)
            ws.write(current, 2, it[2], cell)
            ws.write(current, 3, it[3], cell)
            ws.write(current, 4, it[4], cell)
            ws.write(current, 5, it[5], fmt_color)
            ws.write(current, 6, it[6], num)
            ws.write(current, 7, it[7], num)
            ws.write(current, 8, it[8], num)
            ws.write(current, 9, it[9], num)
            ws.write(current, 10, rank, cell)
            current += 1

        # Итоговая строка
        total_cnt = sum(d['count'] for d in stats.by_developer.values())
        total_a4 = sum(d['a4'] for d in stats.by_developer.values())
        ws.write(current, 0, 'ИТОГО ПО ВСЕМ РАЗРАБОТЧИКАМ:', total_fmt_int)
        ws.write(current, 1, total_cnt, total_fmt_int)
        ws.write(current, 2, total_a4, total_fmt_int)
        ws.write(current, 3, stats.total_errors_cat1, total_fmt_int)
        ws.write(current, 4, stats.total_errors_cat2, total_fmt_int)
        avg1_all = stats.total_errors_cat1 / total_cnt if total_cnt else 0
        avg2_all = stats.total_errors_cat2 / total_cnt if total_cnt else 0
        ws.write(current, 5, avg1_all, total_fmt)
        ws.write(current, 6, avg2_all, total_fmt)
        ws.write(current, 7, stats.total_errors_cat1 / total_a4 if total_a4 else 0, total_fmt)
        ws.write(current, 8, stats.total_errors_cat2 / total_a4 if total_a4 else 0, total_fmt)
        coef_all = (stats.total_errors_cat1 + 0.2 * stats.total_errors_cat2) / total_cnt if total_cnt else 0
        ws.write(current, 9, coef_all, total_fmt)
        ws.write(current, 10, f'Всего: {len(devs)}', total_fmt_int)
        return current + 1

    def _write_additional_info(self, ws, stats, start_row):
        """Дополнительная информация."""
        current = start_row
        bold_title = self.workbook.add_format(FMT_BOLD)
        ws.write(current, 0, 'ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ:', bold_title)
        current += 1
        total_cnt = sum(d['count'] for d in stats.by_developer.values())
        total_a4 = sum(d['a4'] for d in stats.by_developer.values())
        num_devs = len(stats.by_developer)
        if num_devs == 0:
            return current
        coef_all = (stats.total_errors_cat1 + 0.2 * stats.total_errors_cat2) / total_cnt if total_cnt else 0
        ws.write(current, 0, f'• Всего уникальных разработчиков: {num_devs}')
        current += 1
        ws.write(current, 0, f'• Среднее количество документов на разработчика: {total_cnt/num_devs:.2f}')
        current += 1
        ws.write(current, 0, f'• Среднее количество форматов А4 на разработчика: {total_a4/num_devs:.2f}')
        current += 1
        ws.write(current, 0, f'• Общий коэффициент ошибок: {coef_all:.2f}')
        return current + 1

    def _write_quality_distribution(self, ws, stats, start_row):
        """Распределение по качеству (цветное)."""
        current = start_row
        bold_title = self.workbook.add_format(FMT_BOLD)
        ws.write(current, 0, 'РАСПРЕДЕЛЕНИЕ ПО КАЧЕСТВУ РАБОТЫ:', bold_title)
        current += 1

        devs_avg = []
        for dev, d in stats.by_developer.items():
            cnt = d['count']
            if cnt == 0: continue
            avg1 = d['errors1'] / cnt
            devs_avg.append(avg1)

        quality = {'<0.5': 0, '0.5-1.5': 0, '1.5-3': 0, '>=3': 0}
        for avg in devs_avg:
            if avg < THRESHOLD_EXCELLENT:
                quality['<0.5'] += 1
            elif avg < THRESHOLD_GOOD:
                quality['0.5-1.5'] += 1
            elif avg < THRESHOLD_POOR:
                quality['1.5-3'] += 1
            else:
                quality['>=3'] += 1

        fmt_excellent = self.workbook.add_format({'bg_color': COLOR_EXCELLENT, 'border': 1})
        fmt_good = self.workbook.add_format({'bg_color': COLOR_GOOD, 'border': 1})
        fmt_average = self.workbook.add_format({'bg_color': COLOR_AVERAGE, 'border': 1})
        fmt_poor = self.workbook.add_format({'bg_color': COLOR_POOR, 'border': 1})

        mapping = [
            ('<0.5', 'Отличное качество', fmt_excellent),
            ('0.5-1.5', 'Хорошее качество', fmt_good),
            ('1.5-3', 'Среднее качество', fmt_average),
            ('>=3', 'Требует улучшения', fmt_poor)
        ]
        total_devs = len(devs_avg)
        for lbl, desc, fmt in mapping:
            cnt = quality[lbl]
            pct = (cnt / total_devs * 100) if total_devs else 0
            ws.write(current, 0, f'• {desc} ({lbl} ошибок): {cnt} ({pct:.1f}%)', fmt)
            current += 1

        return current

    # -----------------------------------------------------------------
    # Зона 6 – Статистика по месяцам
    # -----------------------------------------------------------------
    def _zone6(self, ws, stats, start_row, hdr, cell, num, section_title_fmt, total_fmt, total_fmt_int):
        current = start_row + 1
        ws.write(current, 0, 'СТАТИСТИКА ПО МЕСЯЦАМ', section_title_fmt)
        current += 2
        headers = ['Месяц','Количество документов','Ошибки кат.1','Ошибки кат.2',
                   'Всего А4','Средн. кол-во ошибок кат.1 на док','Средн. кол-во ошибок кат.2 на док',
                   'Средн. кол-во ошибок кат.1 на А4','Средн. кол-во ошибок кат.2 на А4']
        for c, h in enumerate(headers):
            ws.write(current, c, h, hdr)
        current += 1

        months = sorted(stats.by_month.keys())
        for m in months:
            d = stats.by_month[m]
            cnt = d['count']; e1 = d['errors1']; e2 = d['errors2']; a4 = d['a4']
            avg1 = e1 / cnt if cnt else 0
            avg2 = e2 / cnt if cnt else 0
            a1a4 = e1 / a4 if a4 else 0
            a2a4 = e2 / a4 if a4 else 0
            if avg1 >= THRESHOLD_POOR:
                color = COLOR_POOR
            elif avg1 >= THRESHOLD_GOOD:
                color = COLOR_GOOD
            else:
                color = COLOR_EXCELLENT
            fmt_color = self.workbook.add_format({
                **FMT_CELL_WRAP, 'bg_color': color, 'border': 1, 'num_format': '0.00'
            })
            year, month_num = map(int, m.split('-'))
            month_name = f"{RU_MONTHS[month_num]} {year}"
            ws.write(current, 0, month_name, cell)
            ws.write(current, 1, cnt, cell)
            ws.write(current, 2, e1, cell)
            ws.write(current, 3, e2, cell)
            ws.write(current, 4, a4, cell)
            ws.write(current, 5, avg1, fmt_color)
            ws.write(current, 6, avg2, num)
            ws.write(current, 7, a1a4, num)
            ws.write(current, 8, a2a4, num)
            current += 1

        # Итог
        total_cnt = sum(d['count'] for d in stats.by_month.values())
        total_e1 = stats.total_errors_cat1; total_e2 = stats.total_errors_cat2
        total_a4 = stats.total_a4
        ws.write(current, 0, 'ИТОГО за период:', total_fmt_int)
        ws.write(current, 1, total_cnt, total_fmt_int)
        ws.write(current, 2, total_e1, total_fmt_int)
        ws.write(current, 3, total_e2, total_fmt_int)
        ws.write(current, 4, total_a4, total_fmt_int)
        ws.write(current, 5, total_e1/total_cnt if total_cnt else 0, total_fmt)
        ws.write(current, 6, total_e2/total_cnt if total_cnt else 0, total_fmt)
        ws.write(current, 7, total_e1/total_a4 if total_a4 else 0, total_fmt)
        ws.write(current, 8, total_e2/total_a4 if total_a4 else 0, total_fmt)
        return current + 1

    # -----------------------------------------------------------------
    # Зона 7 – Анализ ошибок
    # -----------------------------------------------------------------
    def _zone7(self, ws, stats, row, bold, cell, num, pct_fmt, section_title_fmt):
        current = row + 1
        ws.write(current, 0, 'АНАЛИЗ ОШИБОК', section_title_fmt)
        current += 2
        ws.write(current, 0, 'Общие показатели:', bold)
        current += 1
        ws.write(current, 0, 'Всего документов:')
        ws.write(current, 1, stats.total_docs)
        current += 1
        ws.write(current, 0, 'Документов с ошибками:')
        ws.write(current, 1, stats.docs_with_errors)
        ws.write(current, 2, stats.docs_with_errors / stats.total_docs if stats.total_docs else 0, pct_fmt)
        current += 1
        ws.write(current, 0, 'Документов без ошибок:')
        ws.write(current, 1, stats.total_docs - stats.docs_with_errors)
        ws.write(current, 2, (stats.total_docs - stats.docs_with_errors) / stats.total_docs if stats.total_docs else 0, pct_fmt)
        current += 1
        ws.write(current, 0, 'Всего форматов А4:')
        ws.write(current, 1, stats.total_a4_errors)
        current += 2
        ws.write(current, 0, 'Статистика ошибок:', bold)
        current += 1
        ws.write(current, 0, 'Всего ошибок категории 1:')
        ws.write(current, 1, stats.total_errors_cat1)
        current += 1
        ws.write(current, 0, 'Всего ошибок категории 2:')
        ws.write(current, 1, stats.total_errors_cat2)
        current += 1
        ws.write(current, 0, 'Максимум ошибок кат.1 в документе:')
        ws.write(current, 1, stats.max_errors_cat1)
        current += 1
        ws.write(current, 0, 'Максимум ошибок кат.2 в документе:')
        ws.write(current, 1, stats.max_errors_cat2)

    # -----------------------------------------------------------------
    # Вставка всех 8 диаграмм
    # -----------------------------------------------------------------
    def _insert_charts(self, ws, stats, num_months, num_devs, num_groups):
        if stats.total_docs == 0:
            return
        chart_row = 8
        row_step = 20

        self._add_pie_errors(ws, chart_row, 1.5, 1.2)
        chart_row += row_step
        self._add_pie_with_errors(ws, chart_row, 1.5, 1.2)
        chart_row += row_step
        self._add_column_months(ws, chart_row, 1.5, 1.2, num_months)
        chart_row += row_step
        self._add_column_errors_months(ws, chart_row, 1.5, 1.2, num_months)
        chart_row += row_step
        self._add_column_dev_docs(ws, chart_row, 1.5, 1.2, num_devs)
        chart_row += row_step
        self._add_column_dev_cat1(ws, chart_row, 1.5, 1.2, num_devs)
        chart_row += row_step
        self._add_column_dev_cat2(ws, chart_row, 1.5, 1.2, num_devs)
        chart_row += row_step
        self._add_column_file_groups(ws, chart_row, 1.5, 1.2, num_groups)

    def _add_pie_errors(self, ws, row, scale_w, scale_h):
        pie1 = self.workbook.add_chart({'type': 'pie'})
        pie1.add_series({
            'name': '=_chart_data!$A$1',
            'categories': '=_chart_data!$A$2:$A$5',
            'values': '=_chart_data!$B$2:$B$5',
            'data_labels': {'value': True, 'category': True, 'separator': '\n', 'num_format': '0'},
            'points': [
                {'fill': {'color': '#FFB3B3'}},
                {'fill': {'color': '#B3D4FF'}},
                {'fill': {'color': '#D4B3FF'}},
                {'fill': {'color': '#B3FFB3'}},
            ]
        })
        pie1.set_title({'name': 'Распределение ошибок'})
        ws.insert_chart(f'L{row}', pie1, {'x_scale': scale_w, 'y_scale': scale_h})

    def _add_pie_with_errors(self, ws, row, scale_w, scale_h):
        pie2 = self.workbook.add_chart({'type': 'pie'})
        pie2.add_series({
            'name': '=_chart_data!$A$7',
            'categories': '=_chart_data!$A$8:$A$9',
            'values': '=_chart_data!$B$8:$B$9',
            'data_labels': {'percentage': True, 'category': True, 'separator': '\n'},
            'points': [
                {'fill': {'color': '#FFB3B3'}},
                {'fill': {'color': '#B3FFB3'}},
            ]
        })
        pie2.set_title({'name': 'Доля документов с ошибками'})
        ws.insert_chart(f'L{row}', pie2, {'x_scale': scale_w, 'y_scale': scale_h})

    def _add_column_months(self, ws, row, scale_w, scale_h, num_months):
        if num_months == 0:
            return
        end_row_months = num_months + 1
        col1 = self.workbook.add_chart({'type': 'column'})
        col1.add_series({
            'name': 'Количество документов',
            'categories': f'=_chart_data!$D$2:$D${end_row_months}',
            'values': f'=_chart_data!$E$2:$E${end_row_months}',
            'data_labels': {'value': True},
            'fill': {'color': '#4472C4'},
        })
        col1.set_title({'name': 'Динамика проверки документов по месяцам'})
        col1.set_x_axis({'name': 'Месяц'})
        col1.set_y_axis({'name': 'Количество документов'})
        col1.set_legend({'position': 'bottom'})
        ws.insert_chart(f'L{row}', col1, {'x_scale': scale_w, 'y_scale': scale_h})

    def _add_column_errors_months(self, ws, row, scale_w, scale_h, num_months):
        if num_months == 0:
            return
        end_row_months = num_months + 1
        col2 = self.workbook.add_chart({'type': 'column'})
        col2.add_series({
            'name': 'Ошибки кат.1',
            'categories': f'=_chart_data!$G$2:$G${end_row_months}',
            'values': f'=_chart_data!$H$2:$H${end_row_months}',
            'fill': {'color': '#C00000'},
            'data_labels': {'value': True, 'num_format': '# ##0;;'},
        })
        col2.add_series({
            'name': 'Ошибки кат.2',
            'categories': f'=_chart_data!$G$2:$G${end_row_months}',
            'values': f'=_chart_data!$I$2:$I${end_row_months}',
            'fill': {'color': '#FF8000'},
            'data_labels': {'value': True, 'num_format': '# ##0;;'},
        })
        col2.set_title({'name': 'Динамика ошибок по месяцам'})
        col2.set_x_axis({'name': 'Месяц'})
        col2.set_y_axis({'name': 'Количество ошибок'})
        col2.set_legend({'position': 'bottom'})
        ws.insert_chart(f'L{row}', col2, {'x_scale': scale_w, 'y_scale': scale_h})

    def _add_column_dev_docs(self, ws, row, scale_w, scale_h, num_devs):
        if num_devs == 0:
            return
        end_row_devs = num_devs + 1
        col3 = self.workbook.add_chart({'type': 'column'})
        col3.add_series({
            'name': 'Количество документов',
            'categories': f'=_chart_data!$K$2:$K${end_row_devs}',
            'values': f'=_chart_data!$L$2:$L${end_row_devs}',
            'fill': {'color': '#4472C4'},
            'data_labels': {'value': True},
        })
        col3.set_title({'name': 'Количество документов по разработчикам'})
        col3.set_x_axis({'name': 'Разработчик'})
        col3.set_y_axis({'name': 'Количество документов'})
        col3.set_legend({'position': 'bottom'})
        ws.insert_chart(f'L{row}', col3, {'x_scale': scale_w, 'y_scale': scale_h})

    def _add_column_dev_cat1(self, ws, row, scale_w, scale_h, num_devs):
        if num_devs == 0:
            return
        end_row_devs = num_devs + 1
        col4 = self.workbook.add_chart({'type': 'column'})
        col4.add_series({
            'name': 'Ошибки кат.1',
            'categories': f'=_chart_data!$K$2:$K${end_row_devs}',
            'values': f'=_chart_data!$M$2:$M${end_row_devs}',
            'fill': {'color': '#C00000'},
            'data_labels': {'value': True},
        })
        col4.set_title({'name': 'Ошибки категории 1 по разработчикам'})
        col4.set_x_axis({'name': 'Разработчик'})
        col4.set_y_axis({'name': 'Количество ошибок'})
        col4.set_legend({'position': 'bottom'})
        ws.insert_chart(f'L{row}', col4, {'x_scale': scale_w, 'y_scale': scale_h})

    def _add_column_dev_cat2(self, ws, row, scale_w, scale_h, num_devs):
        if num_devs == 0:
            return
        end_row_devs = num_devs + 1
        col5 = self.workbook.add_chart({'type': 'column'})
        col5.add_series({
            'name': 'Ошибки кат.2',
            'categories': f'=_chart_data!$K$2:$K${end_row_devs}',
            'values': f'=_chart_data!$N$2:$N${end_row_devs}',
            'fill': {'color': '#FF8000'},
            'data_labels': {'value': True},
        })
        col5.set_title({'name': 'Ошибки категории 2 по разработчикам'})
        col5.set_x_axis({'name': 'Разработчик'})
        col5.set_y_axis({'name': 'Количество ошибок'})
        col5.set_legend({'position': 'bottom'})
        ws.insert_chart(f'L{row}', col5, {'x_scale': scale_w, 'y_scale': scale_h})

    def _add_column_file_groups(self, ws, row, scale_w, scale_h, num_groups):
        if num_groups == 0:
            return
        end_row = num_groups + 1
        chart = self.workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Количество документов',
            'categories': f'=_chart_data!$P$2:$P${end_row}',
            'values': f'=_chart_data!$Q$2:$Q${end_row}',
            'fill': {'color': '#4472C4'},
            'data_labels': {'value': True},
        })
        chart.set_title({'name': 'Количество документов по группам файлов (префикс до "_")'})
        chart.set_x_axis({'name': 'Группа файлов'})
        chart.set_y_axis({'name': 'Количество документов'})
        chart.set_legend({'position': 'bottom'})
        ws.insert_chart(f'L{row}', chart, {'x_scale': scale_w, 'y_scale': scale_h})