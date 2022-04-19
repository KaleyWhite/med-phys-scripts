import pandas as pd
COLORS_PATH = r'T:\Physics\KW\med-phys-spreadsheets\TG-263 Nomenclature with CRMC Colors.xlsm'
def read_colors():
    colors = pd.read_excel(COLORS_PATH, sheet_name='Names & Colors', usecols=['TG-263 Primary Name', 'Color'])
    colors.set_index('TG-263 Primary Name', drop=True, inplace=True)
    return colors['Color']

colors = read_colors()
print(colors['Chestwall_L'])