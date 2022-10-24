import win32com.client
import pandas as pd
import os
import sys
from tqdm import tqdm

work_folder = os.getcwd() + "/"
wdReplaceAll = 2
wdFindContinue = 1
wdExportFormatPDF = 17

df = pd.read_excel(work_folder + "Data.xlsx", engine='openpyxl')
df.columns = [col_name.upper() for col_name in df.columns]

