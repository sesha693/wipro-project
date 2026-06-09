from pathlib import Path
import pandas as pd

fp = Path('tmp_test.xlsx')
with pd.ExcelWriter(fp, engine='openpyxl') as writer:
    df = pd.DataFrame({
        'Account': ['A', 'B'],
        'RU Plan QTR': [10, 20],
        'RU Plan': [5, 15],
        'RU Act': [6, 17],
        'RU Gap': [1, 2],
        'Prev WK GAP': [0, -1],
        'WOW': [2, -3],
        'RU Plan BPM QTR': [12, 18],
        'RU Plan BPM': [4, 14],
        'RU Act BPM': [5, 16],
        'RU BPM Gap': [1, 2],
        'Prev week gap': [0, -2],
        'WOW BPM': [1, -1],
        'Delta Reason': ['reason', ''],
        'Recovery Plan': ['plan', ''],
        'ADH': ['Kumar', 'Prachi'],
    })
    df.to_excel(writer, sheet_name='RU', index=False)
    df2 = df.rename(columns={
        'RU Plan QTR': 'RD Plan QTR',
        'RU Plan': 'RD Plan QTD',
        'RU Act': 'RD Act QTD',
        'RU Gap': 'RD Gap',
        'RU Plan BPM QTR': 'RD Plan BPM QTR',
        'RU Plan BPM': 'RD Plan BPM',
        'RU Act BPM': 'RD Act BPM',
        'RU BPM Gap': 'RD BPM Gap',
    })
    df2.to_excel(writer, sheet_name='RD', index=False)

import sys
sys.path.insert(0, str(Path('.').resolve()))
from src.data_reader import get_all_data

data = get_all_data(str(fp), ['RU', 'RD', 'Netadd'], 'all')
print(data.keys())
print(data['Netadd'][0])
print('OK')
