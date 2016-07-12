# Data Manupulation using Pandas Python


import pandas as pd
import numpy as np

# Load Data into Python memory.


path = file_path
with pd.ExcelFile(path) as xls:
    df = pd.read_excel(xls, Sheet_name )
# Strip headers of all columns
df = df.rename(columns=lambda x: x.strip())

# Cleaning for PT Count.
df['DN / PT#'] = df['DN / PT#'].replace(', ',',')
df['DN / PT#'] = df['DN / PT#'].replace('No D/N on docs','')

# Enter PT Count
df['PT COUNT'] = df['DN / PT#'].apply(lambda x: str(x).count(',')+1 if str(x).count(',') != 0 else (1 if len(str(x))!=0 else 0 ))
