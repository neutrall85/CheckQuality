HEADER_BG = '#C8C8C8'
GREEN_BG = '#C6EFCE'
YELLOW_BG = '#FFEB84'
ORANGE_BG = '#FFC7CE'
RED_BG = '#FF6464'

MONTH_RED_BG = '#FFC7CE'
MONTH_YELLOW_BG = '#FFEB84'
MONTH_GREEN_BG = '#C6EFCE'

# Пороги качественной оценки
THRESHOLD_EXCELLENT = 0.5
THRESHOLD_GOOD = 1.5
THRESHOLD_POOR = 3.0

COLOR_EXCELLENT = GREEN_BG
COLOR_GOOD = YELLOW_BG
COLOR_AVERAGE = ORANGE_BG
COLOR_POOR = RED_BG

FMT_BOLD = {'bold': True}
FMT_HEADER = {
    'bold': True,
    'bg_color': HEADER_BG,
    'border': 1,
    'text_wrap': True,
    'valign': 'vcenter',
    'align': 'center'
}
FMT_CELL_WRAP = {
    'text_wrap': True,
    'valign': 'vcenter',
    'border': 1
}
FMT_PERCENT = {'num_format': '0.00%'}
FMT_NUM_2DEC = {'num_format': '0.00'}