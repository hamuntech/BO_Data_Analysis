import pandas as pd
import re as re

def percentage(part, whole):
  Percentage = round(100 * float(part)/float(whole), 2)
  return str(Percentage) + "%"

def dfs_tabs(df_list, sheet_list, file_name):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0 , startcol=0) #index=False   
    writer.save()

def remove_html_tags(worklog):
    result = re.sub('<.*?>','',worklog)
    return result