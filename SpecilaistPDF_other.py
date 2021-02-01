from pandas import DataFrame
#import dash
import pandas as pd
import numpy as np
import plotly.graph_objs as go
import User_Interface
import webbrowser
import tkinter as tk
#import dash_core_components as dcc
#import dash_html_components as html
#import dash_bootstrap_components as dbc
import datetime
import pymysql
import dash_table
from plotly.subplots import make_subplots
#import reportlab
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter,landscape, A4
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch,cm
from reportlab.lib import utils, styles
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import  TTFont

# --------------------------------------------------------------------------------------------------------------------------
# ------------------------------------------------
# Tkinter used as User Interface just to determine which data is being extracted and the presence of data.
# This UI nothing to do with the automated spool-in take dashboard reports
def alert():
    end_msg = tk.Tk()
    end_msg.title('Report Requested')
    tk.Message(end_msg,text= """The Report has been requested.Click View to see dashboard on browser. 
    Check your given E-Mail for the report
    """).pack()
    tk.Button(end_msg, text = "View Report", command = end_msg.destroy).pack()
    end_msg.geometry('300x200')
    end_msg.mainloop()


def alerttoerror():
    error_msg = tk.Tk()
    error_msg.title('Report Requested')
    tk.Message(error_msg, text=""" 
    Please check Input Data. 
    The data doesn't contain Nominal dancer Values.
    Deviation Position Graph will not be available.   
        """).pack()
    tk.Button(error_msg, text="Proceed", command=error_msg.destroy).pack()
    error_msg.geometry('500x500')
    error_msg.mainloop()



# --------------------------------------------------------------------------
# SQL data acquisition
db = pymysql.connect(db="CustomerDB_706998", user="student",passwd="student",host="86.109.253.36",port=55500)
cursor = db.cursor()
cursor.execute(f"""
SELECT 
Value,`Timestamp`,Alias,BMK,Description
FROM CustomerDB_706998.tProcessData
INNER JOIN CustomerDB_706998.tDatapoint ON CustomerDB_706998.tProcessData.idDatapoint=CustomerDB_706998.tDatapoint.id
INNER JOIN CustomerDB_706998.tMachine ON CustomerDB_706998.tDatapoint.idMachine=CustomerDB_706998.tMachine.id
WHERE `Timestamp` BETWEEN {User_Interface.timefsv} AND {User_Interface.timetsv} 
AND Alias IN ('SpeedAct','SlipwayTemperatureAct','MachineStatus',
'WTC_P14StepsMotor1','WTC_P14StepsMotor10','WTC_P14StepsMotor11','WTC_P14StepsMotor12','WTC_P14StepsMotor13','WTC_P14StepsMotor14','WTC_P14StepsMotor15','WTC_P14StepsMotor16','WTC_P14StepsMotor2','WTC_P14StepsMotor3','WTC_P14StepsMotor4','WTC_P14StepsMotor5','WTC_P14StepsMotor6','WTC_P14StepsMotor7','WTC_P14StepsMotor8','WTC_P14StepsMotor9',
'WTC_P20MaxCurrentMotor1','WTC_P20MaxCurrentMotor10','WTC_P20MaxCurrentMotor11','WTC_P20MaxCurrentMotor12','WTC_P20MaxCurrentMotor13','WTC_P20MaxCurrentMotor14','WTC_P20MaxCurrentMotor15','WTC_P20MaxCurrentMotor16','WTC_P20MaxCurrentMotor2','WTC_P20MaxCurrentMotor3','WTC_P20MaxCurrentMotor4','WTC_P20MaxCurrentMotor5','WTC_P20MaxCurrentMotor6','WTC_P20MaxCurrentMotor7','WTC_P20MaxCurrentMotor8','WTC_P20MaxCurrentMotor9',
'WTC_P34NomValueDancer1','WTC_P34NomValueDancer10','WTC_P34NomValueDancer11','WTC_P34NomValueDancer12','WTC_P34NomValueDancer13','WTC_P34NomValueDancer14','WTC_P34NomValueDancer15','WTC_P34NomValueDancer16','WTC_P34NomValueDancer2','WTC_P34NomValueDancer3','WTC_P34NomValueDancer4','WTC_P34NomValueDancer5','WTC_P34NomValueDancer6','WTC_P34NomValueDancer7','WTC_P34NomValueDancer8','WTC_P34NomValueDancer9',
'WTC_P50RegActive1','WTC_P50RegActive10','WTC_P50RegActive11','WTC_P50RegActive12','WTC_P50RegActive13','WTC_P50RegActive14','WTC_P50RegActive15','WTC_P50RegActive16','WTC_P50RegActive2','WTC_P50RegActive3','WTC_P50RegActive4','WTC_P50RegActive5','WTC_P50RegActive6','WTC_P50RegActive7','WTC_P50RegActive8','WTC_P50RegActive9',
'WTC_P31ActValueDancer1','WTC_P31ActValueDancer10','WTC_P31ActValueDancer11','WTC_P31ActValueDancer12','WTC_P31ActValueDancer13','WTC_P31ActValueDancer14','WTC_P31ActValueDancer15','WTC_P31ActValueDancer16','WTC_P31ActValueDancer2','WTC_P31ActValueDancer3','WTC_P31ActValueDancer4','WTC_P31ActValueDancer5','WTC_P31ActValueDancer6','WTC_P31ActValueDancer7','WTC_P31ActValueDancer8','WTC_P31ActValueDancer9', 
'LubricatingOffTime','LubricatingOnTime','LubricationPulses', 'PowerInput')
AND MCMachine IN ('{User_Interface.mcNsv}')
""")

cursor2 = db.cursor()
cursor2.execute(f"""
SELECT
`Timestamp`,Value,Alias
FROM CustomerDB_706998.tProcessData
INNER JOIN CustomerDB_706998.tDatapoint ON CustomerDB_706998.tProcessData.idDatapoint=CustomerDB_706998.tDatapoint.id
INNER JOIN CustomerDB_706998.tMachine ON CustomerDB_706998.tDatapoint.idMachine=CustomerDB_706998.tMachine.id
AND Alias IN ('LubricatingOffTime')
AND MCMachine IN ('{User_Interface.mcNsv}')
""")

# for Product Code
cursor3 = db.cursor()
cursor3.execute(f"""SELECT 
`Timestamp`, Value
FROM CustomerDB_706998.tProcessDataStrings 
INNER JOIN CustomerDB_706998.tDatapoint ON CustomerDB_706998.tProcessDataStrings.idDatapoint = CustomerDB_706998.tDatapoint.id 
INNER JOIN CustomerDB_706998.tMachine ON CustomerDB_706998.tDatapoint.idMachine=CustomerDB_706998.tMachine.id 
WHERE MCMachine IN ('{User_Interface.mcNsv}') 
""")

MCname = db.cursor()
MCname.execute(f"""
SELECT 
Alias
FROM CustomerDB_999330.tHeadstation
WHERE MCHeadstation IN ('{User_Interface.mcNsv}')""")


results = cursor.fetchall()
results2 = cursor2.fetchall()
results3 = cursor3.fetchall()
MCname_result = DataFrame(data= MCname.fetchall(),columns= ['Alias'])
db.commit()
db.close()


# -----------------------------------------------------------------------------------------
now = datetime.datetime.now()
# -----------------------------------------------------------------------------------------
# Dataframe to extract last Product-code
db_product_code_test = DataFrame(results3, columns=['Timestamp', 'ProductCode'])
product_code = db_product_code_test.iloc[-1]['ProductCode']


# To save the csv file from database as Dataframe
db_data = DataFrame(results, columns= ['Value', 'Timestamp', 'Alias', 'BMK','Description'])



# Pivoting Dataframe for making each parameter an column
machine_data = db_data.pivot(values='Value', index='Timestamp', columns='Alias').reset_index().rename_axis('index', axis=1)
machine_quality = machine_data[['Timestamp', 'SlipwayTemperatureAct','SpeedAct', 'PowerInput']].dropna(thresh=2)
mostrecent_dt = machine_quality['Timestamp'].max()
mostlate_dt = machine_quality['Timestamp'].min()

if db_data['Alias'].str.contains('LubricatingOffTime').any():
    lubricationdb = machine_data[['Timestamp','LubricationPulses', 'LubricatingOffTime']].dropna(thresh=2)
    lubricationdb['Timestamp'] = pd.to_datetime(lubricationdb['Timestamp']).dt.strftime('%d-%b-%y %H:%M:%S')
    Lubr = lubricationdb[['Timestamp','LubricatingOffTime']].dropna()


else:
    lubricationdb = machine_data[['Timestamp', 'LubricationPulses']].dropna(thresh=2)
    test_lub = DataFrame(results2,columns= ['Timestamp','LubricatingOffTime', 'Alias'])
    #test_lub['Timestamp'] = pd.to_datetime(test_lub['Timestamp']).dt.strftime('%d-%b-%y %H:%M:%S')
    frame1 = test_lub.loc[test_lub['Timestamp'] <= mostlate_dt, ['Timestamp','LubricatingOffTime']].tail(3)
    frame2 = DataFrame([[mostlate_dt,test_lub.iloc[-1]['LubricatingOffTime']]], columns= ['Timestamp', 'LubricatingOffTime'])
    Lubr = frame1.append(frame2)
    Lubr['Timestamp'] = pd.to_datetime(Lubr['Timestamp']).dt.strftime('%d-%b-%y %H:%M:%S')

# list of WTC spools
BMK = ['WTC-01', 'WTC-02','WTC-03','WTC-04','WTC-05','WTC-06','WTC-07','WTC-08','WTC-09','WTC-10','WTC-11','WTC-12','WTC-13','WTC-14','WTC-15','WTC-16']

# Machine Status DataFrame
machine_status = machine_data[['Timestamp','MachineStatus']].reset_index(drop=True, inplace=False).dropna()

# P14-StepsMotor all motors(1-16) segregation
p14_unp = db_data.loc[db_data['Description'] == 'P14-StepsMotor', ['Timestamp','BMK','Value']]
P14 = p14_unp.pivot_table(values='Value', index='Timestamp', columns='BMK').reset_index().rename_axis('index', axis=1)


# P20-
p20_unp = db_data.loc[db_data['Description'] == 'P20-MaxCurrentMotor', ['Timestamp','BMK','Value']]
P20 = p20_unp.pivot_table(values='Value', index='Timestamp', columns='BMK').reset_index().rename_axis('index', axis=1)


# P31-
p31_unp = db_data.loc[db_data['Description'] == 'P31-ActValueDancer', ['Timestamp','BMK','Value']]
P31 = p31_unp.pivot_table(values='Value', index='Timestamp', columns='BMK').reset_index().rename_axis('index', axis=1)


# P34-
p34_unp = db_data.loc[db_data['Description'] == 'P34-NomValueDancer', ['Timestamp','BMK','Value']]

# Condition when P34-NomValueDancer is empty. It is not displayed on the Dash
# This sends empty fifth row, if P34 is empty. And a proper Bar Chart if P34 Values are present
if p34_unp.empty is True:
    alerttoerror()
    # DancerDev = go.Figure(data= [], layout= [] )
else:
    P34 = p34_unp.pivot_table(values='Value', index='Timestamp', columns='BMK').reset_index(drop=True).rename_axis(
        'index', axis=1)
    exp = P34.apply(pd.Series.last_valid_index)
    lastvP34 = pd.Series(P34.values[exp, np.arange(P34.shape[1])],
                         index=P34.columns,
                         name='value')
    # Deviation of Dancer position calculation
    try:
      devpos01 = (lastvP34['WTC-01'] - P31['WTC-01'].mean()) / (lastvP34['WTC-01']) * 100
      devpos02 = (lastvP34['WTC-02'] - P31['WTC-02'].mean()) / (lastvP34['WTC-02']) * 100
      devpos03 = (lastvP34['WTC-03'] - P31['WTC-03'].mean()) / (lastvP34['WTC-03']) * 100
      devpos04 = (lastvP34['WTC-04'] - P31['WTC-04'].mean()) / (lastvP34['WTC-04']) * 100
      devpos05 = (lastvP34['WTC-05'] - P31['WTC-05'].mean()) / (lastvP34['WTC-05']) * 100
      devpos06 = (lastvP34['WTC-06'] - P31['WTC-06'].mean()) / (lastvP34['WTC-06']) * 100
      devpos07 = (lastvP34['WTC-07'] - P31['WTC-07'].mean()) / (lastvP34['WTC-07']) * 100
      devpos08 = (lastvP34['WTC-08'] - P31['WTC-08'].mean()) / (lastvP34['WTC-08']) * 100
      devpos09 = (lastvP34['WTC-09'] - P31['WTC-09'].mean()) / (lastvP34['WTC-09']) * 100
      devpos10 = (lastvP34['WTC-10'] - P31['WTC-10'].mean()) / (lastvP34['WTC-10']) * 100
      devpos11 = (lastvP34['WTC-11'] - P31['WTC-11'].mean()) / (lastvP34['WTC-11']) * 100
      devpos12 = (lastvP34['WTC-12'] - P31['WTC-12'].mean()) / (lastvP34['WTC-12']) * 100
      devpos13 = (lastvP34['WTC-13'] - P31['WTC-13'].mean()) / (lastvP34['WTC-13']) * 100
      devpos14 = (lastvP34['WTC-14'] - P31['WTC-14'].mean()) / (lastvP34['WTC-14']) * 100
      devpos15 = (lastvP34['WTC-15'] - P31['WTC-15'].mean()) / (lastvP34['WTC-15']) * 100
      devpos16 = (lastvP34['WTC-16'] - P31['WTC-16'].mean()) / (lastvP34['WTC-16']) * 100
      DevPosy = [devpos01, devpos02, devpos03, devpos04, devpos05, devpos06, devpos07, devpos08, devpos09, devpos10,
                 devpos11, devpos12, devpos13, devpos14, devpos15, devpos16]
      DancerDev = go.Figure(data = [go.Bar(x=BMK, y=DevPosy)],
                           layout = go.Layout(title={'text' : 'Deviation percentage of Dancer Position'},
                                              yaxis={'title': 'Deviation %age'}, autosize= False, width= 1000, height= 500))
      DancerDev.write_image('assets/DancerDev.png')
    except KeyError:
       pass


# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------
# Speed vs Temperature Plot
SpTemp = make_subplots(specs= [[{'secondary_y': True}]])
SpTemp.add_trace(go.Scatter(x = machine_quality['Timestamp'], y = machine_quality['SpeedAct'], name ='Speed Actual',
                            mode = 'lines', connectgaps= True, line_shape = 'hv'), secondary_y= True)


SpTemp.add_trace(go.Scatter(x = machine_quality['Timestamp'], y = machine_quality['SlipwayTemperatureAct'],
                            name = 'Slipway Temperature', mode = 'lines + markers', connectgaps= True, line_shape = 'hv'), secondary_y= False)

SpTemp.add_trace(go.Scatter(x= machine_quality['Timestamp'], y = machine_quality['PowerInput'], name ='Power Input',
                            mode ='lines', opacity= 0.7, connectgaps= True, yaxis='y3', line_shape = 'hv'))

SpTemp.update_layout(title_text ='Speed, Temperature & Power Input', titlefont= {'color': '#0B0B0A'},
                     legend= {'font': {'color':'#0B0B0A'}}, xaxis =  {'color':'#0B0B0A', 'showgrid': True},
                     yaxis= {'color':'#0B0B0A', 'showgrid': True}, yaxis2 = {'color':'#0B0B0A', 'showgrid': True},
                     yaxis3 = {'title': 'Power Input (kW)','color':'#0B0B0A', 'showgrid': True, 'overlaying' : 'y', 'position': 0.99, 'side' :'right'},
                     autosize= False, width= 1300, height= 500, title_font_family = 'Arial')
SpTemp.update_yaxes(title_text='Speed rpm', secondary_y= True)
SpTemp.update_yaxes(title_text='Temperature °C', secondary_y= False)
SpTemp.write_image('assets/SpeedvsTemp.png') # Image stored in assets


# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------
# Steps Motor Line Plot
StepsFig1 = go.Figure()
try:
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14[f'WTC-01'], name=f'P14-WTC01', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(51, 102, 204)'}))
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-02'], name=f'P14-WTC02', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102, 204, 255)'}))
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-03'], name=f'P14-WTC03', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102, 153, 51)'}))
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-04'], name=f'P14-WTC04', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(255,153,0)'}))
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-05'], name=f'P14-WTC05', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(255,255,51)'}))
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-06'], name=f'P14-WTC06', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102,153,204)' }))
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-07'], name=f'P14-WTC07', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102,102,102)'}))
 StepsFig1.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-08'], name=f'P14-WTC08', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(153,204,102)'}))
 StepsFig1.update_layout(title ='Steps Carrier #01 -- #08', yaxis = {'title': 'Steps in Motor'}, autosize= False, width= 1000, height= 500,title_font_family = 'Arial')
except KeyError:
    pass

StepsFig2 = go.Figure()
try:
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-09'], name=f'P14-WTC09', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(51, 102, 204)'}))
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-10'], name=f'P14-WTC10', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102, 204, 255)'}))
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-11'], name=f'P14-WTC11', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102, 153, 51)'}))
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-12'], name=f'P14-WTC12', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(255,153,0)'}))
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-13'], name=f'P14-WTC13', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(255,255,51)'}))
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-14'], name=f'P14-WTC14', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102,153,204)' }))
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-15'], name=f'P14-WTC15', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(102,102,102)'}))
 StepsFig2.add_trace(go.Scatter(x= P14['Timestamp'], y = P14['WTC-16'], name=f'P14-WTC16', mode ='lines', connectgaps = True, marker = {'color' : 'rgb(153,204,102)'}))
 StepsFig2.update_layout(title ='Steps Carrier #09 -- #16', yaxis = {'title': 'Steps in Motor'}, autosize= False, width= 1000, height= 500,title_font_family = 'Arial')
except KeyError:
    pass

StepsFig1.write_image('assets/Steps#8.png')
StepsFig2.write_image('assets/Steps#16.png')


# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------
# Max motor Current plot
MotorCurrentfig1 = go.Figure()
try:
 MotorCurrentfig1.add_trace(go.Box(x= P20[f'WTC-01'],name='WTC01', boxpoints = False, marker = {'color' : 'rgb(51, 102, 204)'}))
 MotorCurrentfig1.add_trace(go.Box(x= P20['WTC-02'], name='WTC02', boxpoints = False, marker = {'color' : 'rgb(102, 204, 255)' }))
 MotorCurrentfig1.add_trace(go.Box(x= P20['WTC-03'], name='WTC03', boxpoints = False, marker = {'color' : 'rgb(102, 153, 51)'}))
 MotorCurrentfig1.add_trace(go.Box(x= P20['WTC-04'], name='WTC04', boxpoints = False, marker = {'color' : 'rgb(255,153,0)'}))
 MotorCurrentfig1.add_trace(go.Box(x= P20['WTC-05'], name='WTC05', boxpoints = False, marker = {'color' : 'rgb(255,255,51)'}))
 MotorCurrentfig1.add_trace(go.Box(x= P20['WTC-06'], name='WTC06', boxpoints = False, marker = {'color' : 'rgb(102,153,204)' }))
 MotorCurrentfig1.add_trace(go.Box(x= P20['WTC-07'], name='WTC07', boxpoints = False, marker = {'color' : 'rgb(102,102,102)'}))
 MotorCurrentfig1.add_trace(go.Box(x= P20['WTC-08'], name='WTC08', boxpoints = False, marker = {'color' : 'rgb(153,204,102)'}))
 MotorCurrentfig1.update_layout(title ='Max Current Motor #01 -- #08', xaxis = {'title': 'Maximum Current Motor'}, autosize= False, width= 1000, height= 500,title_font_family = 'Arial')
except KeyError:
    pass

MotorCurrentfig1.write_image('assets/motorcurrentfig1.png')

MotorCurrentfig2 = go.Figure()
try:
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-09'], name='WTC09', boxpoints = False, marker = {'color' : 'rgb(51, 102, 204)'}))
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-10'], name='WTC10', boxpoints = False, marker = {'color' : 'rgb(102, 204, 255)'}))
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-11'], name='WTC11', boxpoints = False, marker = {'color' : 'rgb(102, 153, 51)'}))
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-12'], name='WTC12', boxpoints = False, marker = {'color' : 'rgb(255,153,0)'}))
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-13'], name='WTC13', boxpoints = False, marker = {'color' : 'rgb(255,255,51)'}))
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-14'], name='WTC14', boxpoints = False, marker = {'color' : 'rgb(102,153,204)' }))
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-15'], name='WTC15', boxpoints = False, marker = {'color' : 'rgb(102,102,102)'}))
 MotorCurrentfig2.add_trace(go.Box(x= P20['WTC-16'], name='WTC16', boxpoints = False, marker = {'color' : 'rgb(153,204,102)'}))
 MotorCurrentfig2.update_layout(title ='Max Current Motor #09 -- #16' , xaxis = {'title': 'Maximum Current Motor'}, autosize= False, width= 1000, height= 500,title_font_family = 'Arial')
except KeyError:
    pass
MotorCurrentfig2.write_image('assets/motorcurrent2.png')


# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------
# PDF design using reportlab-lib
pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
pdfmetrics.registerFont(TTFont('ArialBd', 'ArialBd.ttf'))

logo = 'assets/logo.png'
reportTitle = 'BMV 16 Engineer Report'
font = 'Arial'
fontsize = 10
width, height = A4

# Title and logo
# Used drawString/Image instead of Sampledocuments style because limited number of pages and more images to show.
pdf = canvas.Canvas('PDF-Dashboards/Specialist-706998_PDFReport_test.pdf', pagesize= landscape(letter))
pdf.setFont('ArialBd', 20 )
pdf.drawString(15,570, reportTitle)
pdf.drawImage(logo, 710,540, width= 0.7 * inch, height= 0.7 * inch)

# header line
pdf.setStrokeColor(colors.grey)
pdf.setLineWidth(0.3)
pdf.line(0, 545, 650, 545)
pdf.line(650, 545, 680, 525)
pdf.line(680, 525, 840, 525)

# row two for machine report
pdf.setFont(font, fontsize)
pdf.drawString(15,505,f'Machine Number')
pdf.drawString(160,505,f'{User_Interface.mcNsv}')
pdf.drawString(15,480,f'Material')
pdf.drawString(160,480,f'{product_code}')
pdf.drawString(550, 505, f'Reporting Date')
pdf.drawString(650,505,f"{datetime.datetime.strftime(now, '%d-%m-%Y %H:%M')}")
pdf.setFont(font, fontsize)
pdf.drawString(550,480, f"Shift {datetime.datetime.strftime(mostlate_dt, '%d-%m-%Y %H:%M')}"
                        f"-- {datetime.datetime.strftime(mostrecent_dt, '%d-%m-%Y %H:%M')} ")

# getting images and its pixels for better image
def get_image(x,y,path, width):
    img = utils.ImageReader(path)
    iw, ih = img.getSize()
    aspect = ih / float(iw)
    return pdf.drawImage(x=x, y=y, image= path, width=width, height=(width * aspect), mask= 'auto')


# Pie Images
get_image(x= 5,y= 210,path = 'assets/SpeedvsTemp.png', width= 700)
get_image(x= 10, y= 0, path = 'assets/DancerDev.png', width = 16.4*cm)


# table for Lubricatinf OFF time.
column1Heading = 'TIMESTAMP'
column2Heading = 'LUBRICATING-OFF'
row_array = [column1Heading,column2Heading]
tableHeading = [row_array]
t = Table(data= tableHeading+np.array(Lubr).tolist(), colWidths= 1.7*inch)
t.setStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                ("ALIGN", (0,0), (-1,-1), "CENTER"),
                ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                ('BOX',(0,0,), (-1,-1), 0.25, colors.black)])
t.wrapOn(pdf, width, height )
t.drawOn(pdf,500,70)


# footer line for page1
pdf.setStrokeColor(colors.grey)
pdf.setLineWidth(0.3)
pdf.line(0, 20, 650, 20)
pdf.line(650, 20, 680, 40)
pdf.line(680, 40, 840, 40)
pdf.setFont('ArialBd', 5)
pdf.setFillColor(colors.black)
pdf.drawString(15,10,'Expertise, Customer Driven and Service - in Good Hands with NIEHOFF')

pdf.showPage()


# Second Page
# Title and logo
pdf.setFont('ArialBd', 20 )
pdf.drawString(15,570, reportTitle)
pdf.drawImage(logo, 710,540, width= 0.7 * inch, height= 0.7 * inch)


# header line
pdf.setStrokeColor(colors.grey)
pdf.setLineWidth(0.3)
pdf.line(0, 545, 650, 545)
pdf.line(650, 545, 680, 525)
pdf.line(680, 525, 840, 525)

# row two for machine report
pdf.setFont(font, fontsize)
pdf.drawString(15,505,f'Machine Number')
pdf.drawString(160,505,f'{User_Interface.mcNsv}')
pdf.drawString(15,480,f'Material')
pdf.drawString(160,480,f'{product_code}')
pdf.drawString(550, 505, f'Reporting Date')
pdf.drawString(650,505,f"{datetime.datetime.strftime(now, '%d-%m-%Y %H:%M')}")
pdf.setFont(font, fontsize)
pdf.drawString(550,480, f"Shift {datetime.datetime.strftime(mostlate_dt, '%d-%m-%Y %H:%M')}"
                        f"-- {datetime.datetime.strftime(mostrecent_dt, '%d-%m-%Y %H:%M')} ")


# Graphs on second page
get_image(0,250,'assets/motorcurrentfig1.png',width = 14*cm)
get_image(395,250, 'assets/motorcurrent2.png', width= 14*cm)
get_image(5,40,'assets/Steps#8.png',width= 14*cm)
get_image(400,40, 'assets/Steps#16.png', width=14*cm)

# footer line for page2
pdf.setStrokeColor(colors.grey)
pdf.setLineWidth(0.3)
pdf.line(0, 20, 650, 20)
pdf.line(650, 20, 680, 40)
pdf.line(680, 40, 840, 40)
pdf.setFont('ArialBd', 5)
pdf.setFillColor(colors.black)
pdf.drawString(15,10,'Expertise, Customer Driven and Service - in Good Hands with NIEHOFF')


pdf.save()

# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------
# Excel writer in the project folder
writer = pd.ExcelWriter(f"Specialist_Files/ReportData{User_Interface.mcNsv}-{datetime.datetime.now().strftime('%Y%m%d %H%M%S')}.xlsx",
                        engine= 'xlsxwriter')
machine_quality.to_excel(writer, sheet_name= 'Speed,Temperature & PowerInput')
P14.to_excel(writer, sheet_name= 'StepsCarrier #1-16')
P20.to_excel(writer, sheet_name= 'MaxCurrentMotor #1-16')
P31.to_excel(writer,sheet_name= 'Actual Position Dancer #1-16')
p34_unp.to_excel(writer, sheet_name= 'Nominal Position Dancer ')
machine_status.to_excel(writer, sheet_name = 'Machine Status')
lubricationdb.to_excel(writer, sheet_name = 'Lubrication Data')
writer.save()

# -----------------------------------------------------------------------------------------
# -----------------------------------------------------------------------------------------










