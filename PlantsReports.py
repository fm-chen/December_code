
# coding: utf-8

# In[ ]:


import pandas as pd
import winreg
import numpy
import math

week_input = int(input("input week_num: "))
#PLANT NUMBER
plantlist = ['1b','2','3','4','5','5a','6','7','9','12']
#PLANT CAP -- 9 has no deliveries
cap_df = pd.DataFrame([('1b',10*5),('2',10*5),('3',4*5),('4',4*5),('5',2*5),
                       ('5a',2*5),('6',2*5),('7',2*5),('12',5*5)],columns=('PLANT_NUM', 'Max_Week_Cap'))

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,                          r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
    return winreg.QueryValueEx(key, "Desktop")[0]

#read raw data
file_loc = r"%s\######.xlsx" %(get_desktop())   #raw data file
df = pd.read_excel(file_loc, sheet_name='Schedule ')
df['week_num'] =pd.to_datetime(df['Original ETA']).dt.week
df['PLANT #'] = df['PLANT #'].fillna(0).apply(str)
new = df['PLANT #'].str.split(',', expand=True).fillna(0)
df['split0'] =new[0]
df['split1'] =new[1]
df['split2'] =new[2]
df['split3'] =new[3]

def plant_report(plantnum,weeknum):
    df['PLANT_NUM'] = '%s' %(plantnum)
    report = df.loc[((df['split0']==plantnum)|(df['split1']==plantnum)|
                     (df['split2']==plantnum)|(df['split3']==plantnum)|
                    (df['split0']==plantnum.upper())|(df['split1']==plantnum.upper())|
                     (df['split2']==plantnum.upper())|(df['split3']==plantnum.upper()))
                    &(df['week_num']==weeknum)]
    return report.drop(columns=['split0','split1','split2','split3', 'week_num'])
def obtain_num(plantnum,weeknum):
    pt = plant_report(plantnum,weeknum)
    pt = pd.pivot_table(pt, index = 'PLANT_NUM',values = 'CONTAINER',aggfunc = 'count')
    return pt
def week_report(weeknum):
    res = pd.concat([obtain_num(plantlist[0],weeknum),obtain_num(plantlist[1],weeknum),
               obtain_num(plantlist[2],weeknum),obtain_num(plantlist[3],weeknum),
               obtain_num(plantlist[4],weeknum),obtain_num(plantlist[5],weeknum),
               obtain_num(plantlist[6],weeknum),obtain_num(plantlist[7],weeknum),
              obtain_num(plantlist[8],weeknum),obtain_num(plantlist[9],weeknum)],axis = 0).reset_index()
    res = res.rename(columns={'CONTAINER':'NUM_CONTAINER'})
    res = pd.merge(res,cap_df,how = 'left',on='PLANT_NUM')
    #res['Potential_issue'] = '√'
    #res['Potential_issue'] = res['Potential_issue'].where(res['NUM_CONTAINER']>res['Max_Week_Cap'])
    return res

#write df to excel
file_name = r'%s\PlantReport\Report_Week_%s.xlsx' %(get_desktop(),week_input)
writer = pd.ExcelWriter(file_name,engine='xlsxwriter', date_format='mm/dd/yyyy', datetime_format='mm/dd/yyyy')
#sheet 1 overview report
sheetName = 'week_%s_report' %(week_input)
df_week = week_report(week_input)
df_week.to_excel(writer,sheet_name = sheetName, index = False)
#format 1
workbook  = writer.book
worksheet = writer.sheets[sheetName]
text_format = workbook.add_format({'align': 'center'})
worksheet.set_column('A:C',15,text_format)
# Add a header format.
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'align': 'center',
    'fg_color': '#D7E4BC',
    'border': 1})
#issue flag
issue_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'align': 'center',
    'fg_color': '#FF0000',
    'border': 1})
# Write issue flag with format.
for index, row in df_week.iterrows(): 
    if row[1]>row[2] and math.isnan(row[2])==False:
        worksheet.write(index+1, 1, row[1], issue_format)
#write column heads with format.
for col_num, value in enumerate(df_week.columns.values):
    worksheet.write(0, col_num, value, header_format)
    
plant_list = []
for index, row in df_week.iterrows(): 
    plant_list.append(row[0])

#sheet 2 plant reports
for temp_list in plant_list:
    sheetName = 'plant_%s' %(temp_list)
    df_plant = plant_report(temp_list,week_input).drop(columns=['PLANT_NUM'])
    df_plant.to_excel(writer,sheet_name = sheetName, index = False)
#format 2 
    workbook  = writer.book
    worksheet = writer.sheets[sheetName]
    worksheet.set_column('A:J',15)
    for col_num, value in enumerate(df_plant.columns.values):
        worksheet.write(0, col_num, value, header_format)
#save
writer.save()
print("finished!")

