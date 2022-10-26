import win32com.client
import pandas as pd
import os
import sys
from tqdm import tqdm

print("-------------------------------------------------------------")
print("PWTUDfE: version 2.3")
print("It helps you to Print Word Template Using Data from Excel.")
print("")
print("Use key -p to ptint testp to default printer.")
print("-------------------------------------------------------------")
print("")
print("")
print("Progress:")

work_folder = os.getcwd() + "/"
wdReplaceAll = 2
wdFindContinue = 1
wdExportFormatPDF = 17

data_file = work_folder + "Data.xlsx"
template_file = work_folder + "Template.docx"
if not os.path.isfile(data_file):
    print("There is no file Data.xlsx")
    exit()
if not os.path.isfile(template_file):
    print("There is no file Template.docx")
    exit()

df = pd.read_excel(data_file, engine='openpyxl')
df.columns = [col_name.upper() for col_name in df.columns]

if "STOP" not in df.columns:
    df["STOP"] = "print"
else:
    df["STOP"] = df["STOP"].fillna("print")

if "FILENAME" not in df.columns:
    df["FILENAME"] = df.index
else:
    df["FILENAME"] = df["FILENAME"].fillna("")

df_fnames = df[df["STOP"]=="print"][["FILENAME"]]
    
need_columns = [col_name for col_name in df.columns if col_name[0]=="{" and col_name[-1]=="}" and col_name!="STOP"]
df = df[df["STOP"]=="print"][need_columns]

_oWord = win32com.client.DispatchEx("Word.Application")
_oWord.Visible = True
_oWord.DisplayAlerts = False

for idx in tqdm(df.index):
    _oWord.Documents.Open(template_file)
    for one_element in need_columns:
        one_substitution = list(df[df.index == idx].to_dict()[one_element].values())[0]
        if type(one_substitution)==pd._libs.tslibs.timestamps.Timestamp:
            one_substitution = format(one_substitution, "%d.%m.%Y")
        _oWord.Selection.Find.Execute(one_element, False, False, False, False, False, True, wdFindContinue, False, one_substitution, wdReplaceAll)
    
    if (len(sys.argv) == 2 and sys.argv[1] == "-p"):
        _oWord.ActiveDocument.PrintOut()
    else:
        pdf_file_name = df_fnames[df_fnames.index==idx].values[0][0]
        for one_char in ',.!~<>?/\|*+-&^%$#@"':
            pdf_file_name = pdf_file_name.replace(one_char, "")
            
        pdf_file_name = work_folder + f"{(idx+1):03}. {pdf_file_name}.pdf"
        _oWord.ActiveDocument.ExportAsFixedFormat(OutputFileName:=pdf_file_name, ExportFormat:=wdExportFormatPDF)
    
    _oWord.ActiveDocument.Close(SaveChanges=False)

_oWord.Quit()
