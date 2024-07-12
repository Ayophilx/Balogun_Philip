import pandas as pd
import os
from datetime import date,timedelta
my_year = date.today().year
my_month = date.today().strftime("%B")
import shutil
import sys

if os.path.exists (f"ctyFirstTimer{my_month}{my_year}Backup.xlsx"):
    print("back up available")
    shutil.copyfile(f"ctyFirstTimer{my_month}{my_year}Backup.xlsx",f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx")

elif os.path.exists(f"ctyFirstTimer{my_month}{my_year}.xlsx"):
    shutil.copyfile(f"ctyFirstTimer{my_month}{my_year}.xlsx",f"ctyFirstTimer{my_month}{my_year}Backup.xlsx")

elif os.path.exists(f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx"):
    shutil.copyfile(f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx",f"ctyFirstTimer{my_month}{my_year}Backup.xlsx")

else :
    print("No backups yet")
    
from openpyxl import load_workbook
 
# assign path
path, dirs, files = next(os.walk("/storage/emulated/0/Documents/Pydroid3/__pycache__/"))
  
import xlsxwriter
#creating dataframe.
ctyFirstTimer=pd.DataFrame()
today= date.today()
todayDf=pd.DataFrame()

def firstTimer():
    global firstname,surname,fullname,today,gender,phonenumber,address,ctyFirstTimer,ctydic,ctyDf,todayDf
    firstname = input("Enter firstname: ")
    firstname = firstname.title()#capatilizing firstname
    surname = input("Enter Lastname: ")
    surname = surname.title()#capatilizing surname
    fullname= firstname+" "+surname
    #today= date.today()
    gender = input("Enter gender(male/female): ")
    phonenumber = input("Enter phone number : ")
    phonenumber =str(phonenumber)
    address = input("Enter address: ")
    #creating a dictionary
    
    ctydic={"Name":[fullname],"Gender":[gender],"PhoneNumber":[phonenumber],"Address":[address],"Date":[today]}
    ctyDf =pd.DataFrame(ctydic)
    ctyFirstTimer = pd.concat([ctyFirstTimer,ctyDf],axis = 0,ignore_index=True)
    
    todayDf=ctyFirstTimer
    
   
response = input("Enter total number of entries: ")
response = int(response)
if response==0:
    sys.exit("Entry must be non-zero")
    
    
while True:
    if response!=0:
        firstTimer()
        
    else:
        break
    response=response-1
    
print(ctyFirstTimer)
week_att_analysis = ctyFirstTimer
my_year=today.year
my_month=today.strftime("%B")
print(my_year,my_month,sep="-")



directory=f'ctyFirstTimer{my_month}{my_year}.xlsx'

try:
        if directory not in files:
            cty_object = pd.ExcelWriter(f'ctyFirstTimer{my_month}{my_year}.xlsx',datetime_format ='mmm dd yyyy', date_format ='mmm dd yyyy', engine="xlsxwriter")
            ctyFirstTimer.to_excel(cty_object,sheet_name=f'{my_month}',index=False)
            attendance_analysis=ctyFirstTimer["Name"].groupby(ctyFirstTimer['Gender']).count().add_prefix('count_')
            attendance_analysis =pd.DataFrame(attendance_analysis)
            attendance_analysis.columns=["Count"]
            print("\n\nLINE 58  ATTENDANCE ANALYSIS \n\n",attendance_analysis)
            cty_object.save()
        
        
        
        else:
                            
             df=pd.read_excel(directory,sheet_name=f'{my_month}',usecols=[0,1,2,3,4],dtype={"PhoneNumber":str}).dropna(axis=0,how="all")
             if df.empty:
                 ctyFirstTimer.to_excel(f'ctyFirstTimer{my_month}{my_year}.xlsx', sheet_name=f"{my_month}",index=False)
                 print("\n\nLINE 65 SHOWING df\n\n",df)
                 
             else:
                 ctyFirstTimer=pd.concat([df,ctyFirstTimer],axis=0,ignore_index=True)
             attendance_analysis=ctyFirstTimer["Name"].groupby(ctyFirstTimer['Gender']).count()
             attendance_analysis =pd.DataFrame(attendance_analysis)
             attendance_analysis.columns=["Count"]
             print("\n\nLINE 72  ATTENDANCE ANALYSIS \n\n",attendance_analysis)
        
        cty_object = pd.ExcelWriter(f'ctyFirstTimer{my_month}{my_year}.xlsx',datetime_format ='mmm dd yyyy', date_format ='mmm dd yyyy',mode="w",engine ="xlsxwriter")
        #writing dataframes to excel
        ctyFirstTimer.to_excel(cty_object,sheet_name=f'{my_month}',index=False,header=True)
        attendance_analysis.to_excel(cty_object,sheet_name=f'{my_month}',index=True,startcol=6,header=True)
        ctyFirstTimer["Date"]=pd.to_datetime(ctyFirstTimer.Date)
        print("\n\nLINE 86 SHOWING ctyFirstTimer\n\n",ctyFirstTimer)
        # creating a workbook object
        ctyWorkbook = cty_object.book
        #creating a workbooksheet
        ctyWorksheet=cty_object.sheets[f'{my_month}']
        # set width of the B and C column
        ctyWorksheet.set_column('A:F', 15)
        # here we create a format object for header
        header_format_object = ctyWorkbook.add_format({ 'bold': True, 'italic' : True,'text_wrap': True,'valign': 'top','font_color': 'blue', 'border': 2})
        # Write the column headers with the defined format.
        for col_number, value in enumerate(ctyFirstTimer.columns.values):
            ctyWorksheet.write(0, col_number,value,header_format_object)# syntax for .write(row, column [ , token [ , format ] ])
            ctyWorksheet.write("G4","Total_count",header_format_object)
            # Using write_formula to calculate total attendance
            ctyWorksheet.write_formula('H4', '{=SUM(H2, H3)}')
      #for col_number, value in enumerate(attendance_analysis.columns.values):
        #    ctyWorksheet.write(0, col_number+5,value,header_format_object)# syntax for .write(row, column [ , token [ , format ] ])
        print(ctyFirstTimer.columns.values)
        # creat chart object
        chartType= ctyWorkbook.add_chart({"type":"doughnut"})
        #add series to chart
        chartType.add_series({"name":f"={my_month}!G1","categories":f"={my_month}!$G$2:$G$3","values":f"={my_month}!$H$2:$H$3",'data_labels': {'value': 1,"category":1},})
        # Add a chart title 
        chartType.set_title ({'name': 'Attendance by Gender'})
        # Set an Excel chart style.
        chartType.set_style(11)
        ctyWorksheet.insert_chart("G6",chartType)
        ctyWorksheet.activate()
        cty_object.save()
        #prev_date=today - timedelta(days = 1)
        
                              
except:
    print("error occured\n\n")
    shutil.copyfile(f"ctyFirstTimer{my_year}Backup.xlsx",f"ctyFirstTimer{my_month}{my_year}.xlsx")
    
    df=pd.read_excel(f"ctyFirstTimer{my_month}{my_year}Backup.xlsx",sheet_name=f'{my_month}',usecols=[0,1,2,3,4],dtype={"PhoneNumber":str}).dropna(axis=0,how="all")
    print("\n\nLINE 65 SHOWING df\n\n",df)
    ctyFirstTimer=pd.concat([df,ctyFirstTimer],axis=0,ignore_index=True)
    attendance_analysis=ctyFirstTimer["Name"].groupby(ctyFirstTimer['Gender']).count()
    attendance_analysis =pd.DataFrame(attendance_analysis)
    attendance_analysis.columns=["Count"]
    print("\n\nLINE 72  ATTENDANCE ANALYSIS \n\n",attendance_analysis)
    
    ctyFirstTimer["Date"]=pd.to_datetime(ctyFirstTimer.Date)
    print("\n\nLINE 86 SHOWING ctyFirstTimer\n\n",ctyFirstTimer)
    # creating a workbook object
    ctyWorkbook = cty_object.book
    #creating a workbooksheet
    ctyWorksheet=cty_object.sheets[f'{my_month}']
    # set width of the B and C column
    ctyWorksheet.set_column('A:F', 15)
    # here we create a format object for header.   
    header_format_object = ctyWorkbook.add_format({ 'bold': True, 'italic' : True,'text_wrap': True,'valign': 'top','font_color': 'blue', 'border': 2})
    # Write the column headers with the defined format.
    for col_number, value in enumerate(ctyFirstTimer.columns.values):
        ctyWorksheet.write(0, col_number,value,header_format_object)# syntax for .write(row, column [ , token [ , format ] ])
        ctyWorksheet.write("G4","Total_count",header_format_object)
        # Using write_formula to calculate total attendance
        ctyWorksheet.write_formula('H4', '{=SUM(H2, H3)}')
    #for col_number, value in enumerate(attendance_analysis.columns.values):
        #    ctyWorksheet.write(0, col_number+5,value,header_format_object)# syntax for .write(row, column [ , token [ , format ] ])
    print(ctyFirstTimer.columns.values)
     # creat chart object
    chartType= ctyWorkbook.add_chart({"type":"doughnut"})
     #add series to chart
    chartType.add_series({"name":f"={my_month}!G1","categories":f"={my_month}!$G$2:$G$3","values":f"={my_month}!$H$2:$H$3",'data_labels': {'value': 1,"category":1},})
     # Add a chart title 
    chartType.set_title ({'name': 'Attendance by Gender'})
     # Set an Excel chart style.
    chartType.set_style(11)
    ctyWorksheet.insert_chart("G6",chartType)
    ctyWorksheet.activate()
    cty_object.save()
     #prev_date=today - timedelta(days = 1)


myPath=f"{today}attendance_.xlsx"
if myPath not in files:
         todayAttObj = pd.ExcelWriter(myPath,datetime_format ='mmm dd yyyy', date_format ='mmm dd yyyy', engine="xlsxwriter")
         todayDf.to_excel(todayAttObj,sheet_name=f'{today}',index=False)
    #todayAttObj.save()
else:
        df=pd.read_excel(myPath,sheet_name=f'{today}',usecols=[0,1,2,3,4],dtype={"PhoneNumber":str})
        print("\n\nLINE 177 SHOWING Today's Attendance\n\n",df)
        today_attend=pd.concat([df,todayDf],axis=0,ignore_index=True)
        todayAttObj = pd.ExcelWriter(myPath,datetime_format ='mmm dd yyyy', date_format ='mmm dd yyyy', engine="xlsxwriter")
        today_attend.to_excel(todayAttObj,sheet_name=f'{today}',index=False)
# creating a workbook object
attendance_Workbook = todayAttObj.book
#creating a workbooksheet
attendance_Worksheet=todayAttObj.sheets[f'{today}']
# set width of the B and C column
attendance_Worksheet.set_column('A:F', 15)
attendance_Worksheet.activate()
todayAttObj.save()
n=1
while n<=7 :
       x=date.today()-timedelta(days=n)
       #print("\n\n",x,"\n\n")
       prev_path=f"{x}attendance.xlsx"
       if os.path.exists(prev_path):
           os.remove(prev_path)
       else:
        print("path not found")
        n+=1
    
 
 
#if os.path.exists (f"ctyFirstTimer{my_month}{my_year}Backup.xlsx"):
#    print("back up available")
#else:
if os.path.exists(f"ctyFirstTimer{my_month}{my_year}.xlsx"):
                shutil.copyfile(f"ctyFirstTimer{my_month}{my_year}.xlsx",f"ctyFirstTimer{my_month}{my_year}Backup.xlsx")
                shutil.copyfile(f"ctyFirstTimer{my_month}{my_year}.xlsx",f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx")
else:
    print("Path not found")   
    
    

df1=pd.read_excel(f"ctyFirstTimer{my_month}{my_year}Backup.xlsx",sheet_name=f'{my_month}',usecols=[0,1,2,3,4],dtype={"PhoneNumber":str}).dropna(axis=0,how="all")

print(df1)
print("\n\n",len(df1))

df2=pd.read_excel(f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx",usecols=[0,1,2,3,4],dtype={"PhoneNumber":str}).dropna(axis=0,how="all")

print(df2)
print("\n\n",len(df2))

if len(df1)>len(df2):
    shutil.copyfile(f"ctyFirstTimer{my_month}{my_year}Backup.xlsx",f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx")
    
else:
    shutil.copyfile(f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx",f"ctyFirstTimer{my_month}{my_year}Backup.xlsx")
    
    
    
df2=pd.read_excel(f"/storage/emulated/0/Documents/Pydroid3/__pycache__/ctybackup/ctyFirstTimer{my_month}{my_year}Backup2.xlsx",usecols=[0,1,2,3,4],dtype={"PhoneNumber":str}).dropna(axis=0,how="all")

print(df2)
print("\n\n",len(df2)) 