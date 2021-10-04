import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
sns.set()

from utility.functions import percentage
from utility.functions import dfs_tabs
from utility.functions import remove_html_tags

import re

pd.options.mode.chained_assignment = None

#READ EXCEL FILE
all_tickets = pd.read_excel('./Files/AutoResV2.xlsx', sheet_name='AUTO_RES_ALL')
auto_res_tickets= all_tickets.dropna()

all_error_tickets = pd.read_excel('./Files/AutoResV2.xlsx', sheet_name='ERRORS')
error_tickets= all_error_tickets.dropna()

error_tickets['Create Date'] = error_tickets['Create Date'].dt.strftime('%m-%d-%Y')

#Count unique Tickets
Unique_Ticket_Count = auto_res_tickets["Ticket Id"].nunique()
Unique_Tickets_with_Errors = error_tickets["Ticket Id"].nunique()

#Only keep unique Tickets
unique_auto_res_df = auto_res_tickets.drop_duplicates(subset='Ticket Id')
unique_error_tic_df = error_tickets.drop_duplicates(subset='Ticket Id')

#Subtract unique Tickets from total number of Tickets
#Difference = Unique_Ticket_Count - Unique_Tickets_with_Errors
#Success_Rate = percentage(Difference, Unique_Ticket_Count)

grouped = unique_error_tic_df.groupby(['Error Type','Automation Bot Details'])['Ticket Id'].count()
grouped = grouped.to_frame().dropna()

grouped_by_bot = unique_error_tic_df.groupby(['Automation Bot Details'])['Ticket Id'].count()
grouped_by_bot = grouped_by_bot.to_frame().dropna()

grouped_by_Log = unique_error_tic_df.groupby(['Log Type'])['Ticket Id'].count()
grouped_by_Log = grouped_by_Log.to_frame().dropna()

gropued_by_error_log = unique_error_tic_df.groupby(['Error Type','Log Type'])['Ticket Id'].count()

grouped_by_error = unique_error_tic_df.groupby(['Error Type'])['Ticket Id'].count()
grouped_by_error = grouped_by_error.dropna()


min_date=min(unique_error_tic_df['Create Date'])
max_date=max(unique_error_tic_df['Create Date'])

#graph_df = unique_auto_res_df.groupby('Create Date').apply(lambda x: pd.Series({'Total_count': len(x),'Auto Resolution Triggered': (x['Error Type'] == 'MANUAL INVESTIGATION REQUIRED').sum()}))
counts_df=pd.DataFrame()
counts_df['Total Count'] = unique_auto_res_df.groupby(pd.Grouper(key='Create Date',freq='1D'))['Ticket Id'].count()
unique_error_tic_df['Create Date'] = pd.to_datetime(unique_error_tic_df['Create Date'])
counts_df['Error Count'] = unique_error_tic_df.groupby(pd.Grouper(key='Create Date',freq='1D'))['Error Type'].count()

#count_by_error = unique_error_tic_df.groupby(pd.Grouper(key='Create Date',freq='1D'))['Error Type'].count()

##################################################
#SUCCESS CALC
#auto_res_log = auto_res_tickets['Log Type'] == "Automated Resolution"
auto_val_log = auto_res_tickets['Log Type'] == "Automated Validation"
matching_string = auto_res_tickets['Worklog Details'] == "MANUAL INVESTIGATION REQUIRED"

#SUCCESSFUL
successful_val = auto_res_tickets[auto_val_log & matching_string].count()

if not(successful_val.any()):
    successful_val = 0

##################################################
#DATA PRINT STATEMENTS

#print("Auto Resolution Stats from: ", min_date, "to: ", max_date)
#print("Total Tickets Received: ", Unique_Ticket_Count)
#print("Condition Not Met:", Unique_Tickets_with_Errors)
#print("Successful:", Unique_Ticket_Count - Unique_Tickets_with_Errors)
#print("Success Rate: ", Success_Rate)
#print("=" * 30)
print("GROUPED BY ERROR AND AUTO: \n", grouped)
print("=" * 30)
print("GROUPED BY BOT: \n",grouped_by_bot)
#print("=" * 30)
#print("GROUPED BY LOG TYPE: \n", grouped_by_Log)
print("=" * 30)
#print("GROUPED BY ERROR & LOG TYPE: \n", gropued_by_error_log)
#print("=" * 30)
#print("GROUPED BY ERROR: \n" , grouped_by_error)
#print("successful_val ", successful_val)

print("total_count  ", counts_df)
##################################################

#EMIT TO EXCEL

#stats_df = pd.DataFrame({'Total Tickets Received': Unique_Ticket_Count, 'Condition Not Met': Unique_Tickets_with_Errors, 'Successful': successful_val}, index=[0])

Dataframes = [grouped, grouped_by_bot, counts_df, unique_auto_res_df, unique_error_tic_df]
Sheets = ['Grouped', 'Grouped By Bot', 'TimeSeries', 'All Tickets', 'Tickets with Errors']

#Call function
dfs_tabs(Dataframes, Sheets, 'AutoResV2.xlsx')




