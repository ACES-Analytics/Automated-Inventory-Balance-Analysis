# -*- coding: UTF-8 -*-
"""
This script is to analyze aging records
Created on Thu Dec 15 07:36:00 2022
@author ACES ANALYTICS TEAM

"""

"""
Index
## 1. Create start time of running script
## 2. Import moduels, packages
## 3. Define date, path, file names
## 4. Import raw data
## 5. Manipulate raw data
## 6. Analyze aging of inventory balance
      - output file: ageing_report_ot
## 7. Inventory balance analysis by ageing
      - output file: ageing_sum_ot
## 8. Inventory balance analysis by profit center
      - output file: prft_sum_ot
## 9. Inventory analysis by ageing and profit center
      Output files:
      - ageing_prft_cost_ot         
      - ageing_prft_volume_ot 
## 10. Plots
       Output plots:
       - Nested pie plot
       - Donut plot
## 11. Format worksheets with visualized plots
## 12. Export worksheets to excel
"""

"""
## 1. Create start time of running script
"""
# Create start time of script running
from timeit import default_timer as timer
start = timer()

"""
## 2. Import modules, packages
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Qt5Agg")  # Do this before importing pyplot.
from matplotlib import pyplot as plt
from millify import millify

"""
## 3. Define date, path, file names
"""

# Define closing date
import datetime
closing_date = datetime.datetime(2022, 12, 31)
closing_date_str = closing_date.strftime("%d-%b-%y")

# Define path
path_from = r'D:\Python\ACES_Analytics\ACES005\Input'
path_to = r'D:\Python\ACES_Analytics\ACES005\Output'

# Define file names
input_file = path_from + "\\" + "ACES001_Inventory Records.xlsx"
output_file = path_to + "\\" + "ACES005_Inventory Aging Report _ " + closing_date_str + ".xlsx"

"""
## 4. Import raw data 
"""
raw_data = pd.read_excel(input_file)
#fixed_ageing_bucket = pd.read_excel(input_file2)

"""
## 5. Manipulate raw data
"""

# Fill zero to na cell
raw_data = raw_data.fillna(0) #method 2: raw_data =raw_data.replace(np.nan,0)

# Calculate lenth of raw data
rec_len = len(raw_data)

# Calculate sum of debit and credit
ttl_Volume = raw_data['Volume\n (kg)'].sum()
ttl_cost = raw_data['Inventory Cost\n ($)'].sum()

"""
## 6. Analyze aging of inventory balance
"""
# Calculate days between posting date and closing date
raw_data['Date'] = pd.to_datetime(raw_data['Date'])
raw_data['Ageing Days'] = raw_data['Date'].apply(lambda x: closing_date - x)

# Convert days into integer type
raw_data['Ageing Days'] = raw_data['Ageing Days'].dt.days

# Define ageing bucket
Ageing_Condition = [
(raw_data['Ageing Days'] >= 0) & (raw_data['Ageing Days'] <= 5),
(raw_data['Ageing Days'] > 5) & (raw_data['Ageing Days'] <= 10),
(raw_data['Ageing Days'] > 10) & (raw_data['Ageing Days'] <= 15),
(raw_data['Ageing Days'] > 15) & (raw_data['Ageing Days'] <= 20),
(raw_data['Ageing Days'] > 20) & (raw_data['Ageing Days'] <= 25),
(raw_data['Ageing Days'] > 25) & (raw_data['Ageing Days'] <= 30),
(raw_data['Ageing Days'] > 30)]

Ageing_Categories = ["0-5 Days", "5-10 Days", "10-15 Days", "15-20 Days", "20-25 Days", "25-30 Days",">30 Days" ]

# Add one new column to calculate  Ageing Bucket
raw_data['Ageing Bucket'] = np.select(Ageing_Condition,Ageing_Categories)

# Selecting columns for ageing report
ageing_report = raw_data.loc[0:,['Date', 'Day', 'Profit Center', 'Mtl Grp', 'Mtl Code', 'Mtl Text',
       'Volume\n (kg)', 'Inventory Cost\n ($)','Ageing Days', 'Ageing Bucket']]

# Change date format as DDMMYYYY
ageing_report['Date'] = pd.to_datetime(ageing_report['Date'])
ageing_report ['Date'] = (ageing_report['Date'].apply(lambda x: x.strftime("%d-%b-%y")))
ageing_report['Date'] = "'" + ageing_report['Date']

# Generate dataframe to get total of columns
# Use DataFrame.loc[] and pandas.Series() to get total of columns
ttl = pd.DataFrame({"Volume\n (kg)": ttl_Volume,
                    "Inventory Cost\n ($)": ttl_cost,
                    "Profit Center": "Total"
                    },index = ['Ttl'])

# Create a new dataframe for output purpose
ageing_report1 = ageing_report.copy().sort_values(by="Date")
ageing_report_ot = pd.concat([ageing_report1,ttl])
ageing_report_ot = ageing_report_ot.fillna(" ")

""" 
## 7. Inventory balance analysis by ageing
"""
# Create dataframe with complete ageing bucket
ageing_bucket_o = pd.DataFrame({"0-5 Days": 0,
                                "5-10 Days": 0,
                                "10-15 Days": 0,
                                "15-20 Days": 0,
                                "20-25 Days": 0,
                                "25-30 Days": 0,
                                ">30 Days": 0
                                },index = ['Cost'])

ageing_sum = pd.pivot_table(ageing_report1,values = ['Volume\n (kg)','Inventory Cost\n ($)',],columns=['Ageing Bucket'],aggfunc = np.sum)
ageing_sum_ot = ageing_sum.copy()
ageing_sum_ot = pd.concat([ageing_bucket_o, ageing_sum_ot])
nan_value1 = float("NaN")
ageing_sum_ot.replace(0, nan_value1, inplace=True)
ageing_sum_ot.dropna(how='all', axis=0, inplace=True)
ageing_sum_ot = ageing_sum_ot.fillna(0)
ageing_sum_ot['Total'] =  ageing_sum_ot.sum(axis=1)
#ageing_sum_ot.loc['Inventory Cost\n ($)',:] = ageing_sum_ot.loc['Inventory Cost\n ($)',:].apply(lambda x: millify(x,precision = 2))

"""
8. Inventory balance analysis by profit center
"""
prft_sum = pd.pivot_table(raw_data,values = ['Inventory Cost\n ($)','Volume\n (kg)'],
                            index = ['Profit Center'], aggfunc = np.sum)

prft_sum1 = prft_sum.copy().sort_values(by="Inventory Cost\n ($)", ascending=False)
prft_sum2 = prft_sum1.T
prft_sum2['Total'] = prft_sum2.sum(axis=1)

# Sort columns in defined order
col_cats = pd.Categorical(prft_sum2.columns.get_level_values(0),
    categories=['Veg-i'] + ['Veg-ii'] + ['Veg-iii'] + ['Veg-iv'] + ['Fruit'] + ['Others']+ ['Total'],
                                      ordered=True)

prft_sum2.columns = col_cats
prft_sum2.sort_index(axis=1,inplace=True)
prft_sum_ot = prft_sum2.copy()
#prft_sum_ot.loc['Inventory Cost\n ($)',:] = prft_sum_ot.loc['Inventory Cost\n ($)',:].apply(lambda x: millify(x,precision = 2))

"""
## 9. Inventory analysis by ageing and profit center
"""
## Analysis in cost
ageing_prft_cost = pd.pivot_table(raw_data,values = ['Inventory Cost\n ($)'],
                                  index = ('Ageing Bucket'),columns = ['Profit Center'],
                                  aggfunc=np.sum, margins = True, margins_name='Total')

ageing_prft_cost .columns.names = (['Inventory Cost\n ($)', 'Profit Center'])

# Sort index with specific orders
ageing_prft_cost1 = ageing_prft_cost .reset_index()

ageing_cats = pd.Categorical(ageing_prft_cost1['Ageing Bucket'],
                             categories=['0-5 Days'] + ['5-10 Days'] + ['10-15 Days'] + ['15-20 Days'] + ['20-25 Days']
                                        + ['25-30 Days'] + ['> 30 Days'] + ['Total'], ordered=True)

ageing_prft_cost1['Ageing Bucket'] = ageing_cats

ageing_prft_cost1.set_index(['Ageing Bucket'], inplace=True)

ageing_prft_cost1 = ageing_prft_cost1.sort_index(level=['Ageing Bucket'])

# Sort Columns
clmn_cats = pd.Categorical(ageing_prft_cost1.columns.get_level_values(1),
    categories=['Veg-i'] + ['Veg-ii'] + ['Veg-iii'] + ['Veg-iv'] + ['Fruit'] + ['Others']+ ['Total'],
                                      ordered=True)
clmn_lev0_cats = ['Inventory Cost\n ($)']

ageing_prft_cost1.columns  = pd.MultiIndex.from_product([clmn_lev0_cats, clmn_cats],
                           names=['Inventory Cost\n ($)', 'Profit Center'])

ageing_prft_cost1.sort_index(axis=1,inplace = True)
ageing_prft_cost_ot = ageing_prft_cost1.copy()

## In volume
ageing_prft_volume = pd.pivot_table(raw_data,values = ['Volume\n (kg)'],
                                  index = ('Ageing Bucket'),columns = ['Profit Center'],
                                  aggfunc=np.sum, margins = True, margins_name='Total')

ageing_prft_volume.columns.names = (['Volume\n (kg)', 'Profit Center'])

# Sort index with specific orders
ageing_prft_volume1 = ageing_prft_volume .reset_index()

ageing_cats = pd.Categorical(ageing_prft_volume1['Ageing Bucket'],
                             categories=['0-5 Days'] + ['5-10 Days'] + ['10-15 Days'] + ['15-20 Days'] + ['20-25 Days']
                                        + ['25-30 Days'] + ['> 30 Days'] + ['Total'], ordered=True)

ageing_prft_volume1['Ageing Bucket'] = ageing_cats

ageing_prft_volume1.set_index(['Ageing Bucket'], inplace=True)

ageing_prft_volume1 = ageing_prft_volume1.sort_index(level=['Ageing Bucket'])

# Sort Columns
clmn_cats = pd.Categorical(ageing_prft_volume1.columns.get_level_values(1),
    categories=['Veg-i'] + ['Veg-ii'] + ['Veg-iii'] + ['Veg-iv'] + ['Fruit'] + ['Others']+ ['Total'],
                                      ordered=True)
clmn_lev0_cats = ['Volume\n (kg)']

ageing_prft_volume1.columns  = pd.MultiIndex.from_product([clmn_lev0_cats, clmn_cats],
                           names=['Volume\n (kg)', 'Profit Center'])

ageing_prft_volume1.sort_index(axis=1,inplace = True)
ageing_prft_volume_ot = ageing_prft_volume1.copy()

"""
## 10. Plots
"""

##10.1 Donut plot
# Create dataset
prft_forplot = prft_sum2.drop("Total", axis = 1)
prft_forplot = round((prft_forplot / 1000000),2)
amt_plt = prft_forplot.loc['Inventory Cost\n ($)',:].tolist()
prft_plt = prft_forplot.columns.tolist()

font_color = '#525252'

# Create colors
a,b,c,d,e,f,g,h = [plt.cm.winter,plt.cm.spring,plt.cm.gist_heat, plt.cm.pink, plt.cm.summer,  plt.cm.autumn, plt.cm.copper, plt.cm.cool,]

colors = [a(.7), b(.7), c(.7), d(.7), e(.7), f(.7), g(.7), h(.7)]

# Wedge properties
wp = { 'linewidth' : 1, 'edgecolor': "white"}

def func(pct, allvalues):
    absolute = (pct / 100 * np.sum(allvalues))
    return "{:.1f}%\n({:,.1f}m $)".format(pct,absolute)

# Create plot
fig1, ax1 = plt.subplots(figsize = (6,6))
wedges, texts, autotexts = ax1.pie(amt_plt,
                                  autopct = lambda pct: func(pct, amt_plt),
                                  labels = prft_plt,
                                  shadow = False,
                                  colors = colors,
                                  startangle = 90,
                                  wedgeprops =wp,
                                  pctdistance = 0.85,
                                  textprops = dict(color = font_color,size = 10, font = 'Verdana'))  #ff10e4 #ffa500 #030803 #weight = "bold"

#plt.setp(autotexts, size =5, color = font_color)
#ax.set_title(" ")

# Set background as transparent
fig1.patch.set_facecolor('none')

# draw circle
centre_circle = plt.Circle((0, 0), 0.60, fc='white')
fig1 = plt.gcf()

# Adding Circle in Pie chart
fig1.gca().add_artist(centre_circle)

# Adding Title of chart
plt.title('By Profit Center',fontsize = 12, color = font_color,font = 'Verdana')

# Add text watermark
fig1.text(0.51, 0.5, 'ACES Analytics Team', fontsize = 16, color = '#fed1f5', ha = 'center', va = 'top',
          alpha = 0.3)

# Displaying Chart
plt.show()

## 10.2 Nested Pie Plot
# Edit data for pie plot
# Create data for nested pie plot
ageing_prft_plt = pd.pivot_table(raw_data,values = ['Inventory Cost\n ($)'],
                                  index = ('Ageing Bucket','Profit Center'),aggfunc=np.sum)

# Sort index with specific orders
ageing_prft_plt1 = ageing_prft_plt.reset_index()

ageing_cats1 = pd.Categorical(ageing_prft_plt1['Ageing Bucket'],
                             categories=['0-5 Days'] + ['5-10 Days'] + ['10-15 Days'] + ['15-20 Days'] + ['20-25 Days']
                                        + ['25-30 Days'] + ['> 30 Days'] , ordered=True)

prft_cats1 = pd.Categorical(ageing_prft_plt1['Profit Center'],
                             categories=['Veg-i'] + ['Veg-ii'] + ['Veg-iii'] + ['Veg-iv'] + ['Fruit'] + ['Others'],
                                      ordered=True)

ageing_prft_plt1['Ageing Bucket'] = ageing_cats1
ageing_prft_plt1['Profit Center'] = prft_cats1
ageing_prft_plt1.set_index(['Ageing Bucket','Profit Center'], inplace=True)
ageing_prft_plt2 = ageing_prft_plt1.sort_index(level=['Ageing Bucket','Profit Center'])

vals = round((ageing_prft_plt2['Inventory Cost\n ($)'] / 1000000), 6)
#vals = prft_vals.loc['Inventory Cost\n ($)',: ].tolist()

ageing_forplot = ageing_sum.copy()
ageing_forplot2 = pd.concat([ageing_bucket_o, ageing_forplot])
nan_value = float("NaN")
ageing_forplot2.replace(0, nan_value, inplace=True)
ageing_forplot2.dropna(how='all', axis=1, inplace=True)
ageing_forplot2.dropna(how='all', axis=0, inplace=True)

# Create dataset
ageing_forplot3 = round((ageing_forplot2 / 1000000),6)
amt_plt = ageing_forplot3.loc['Inventory Cost\n ($)',: ].tolist()
ageing_bucket_plt = ageing_forplot3.columns.tolist()

facecolor = 'none' #eaeaf2
font_color = '#525252'
hfont = {'fontname':'Verdana'}
labels = ageing_bucket_plt
size = 0.3

# Major category values = sum of minor category values
group_sum = amt_plt

# Create a figure
fig2, ax2 = plt.subplots(figsize=(6,6), facecolor=facecolor)

# Create colors
a,b,c,d,e,f,g,h = [plt.cm.winter, plt.cm.cool,plt.cm.gist_heat, plt.cm.pink, plt.cm.summer,  plt.cm.autumn, plt.cm.copper,plt.cm.spring]

outer_colors = [a(.7), b(.7), c(.7), d(.7), e(.7), f(.7), g(.7), h(.7)]
inner_colors = [a(.7), a(.6), a(.5), a(.4), a(.3), a(.2),
                b(.7), b(.6), b(.5), b(.4), b(.3), b(.2),
                c(.7), c(.6), c(.5), c(.4), c(.3), c(.2),
                d(.7), d(.6), d(.5), d(.4), d(.3), d(.2),
                e(.7), e(.6), e(.5), e(.4), e(.3), e(.2),
                f(.7), f(.6), f(.5), f(.4), f(.3), f(.2),
                g(.7), g(.6), g(.5), g(.4), g(.3), g(.2),
                h(.7), h(.6), h(.5), h(.4), h(.3), h(.2)]

# Draw plot

def func(pct, allvalues):
    absolute = (pct / 100.*np.sum(allvalues))
    return "{:.1f}%\n({:,.1f}m $)".format(pct,absolute)

ax2.pie(group_sum,
       autopct = lambda pct: func(pct, group_sum),
       pctdistance = 0.85,
       radius=1,
       colors=outer_colors,
       labels=labels,
       textprops={'color':font_color,'size' : 10},
       wedgeprops=dict(width=size, edgecolor='w'))

ax2.pie(vals,
       radius=1-size, # size=0.3
       colors=inner_colors,
       wedgeprops=dict(width=size, edgecolor='w'))

# set a title
ax2.set_title('By Ageing',fontsize = 12, color = font_color,font = 'Verdana')

# Add text watermark
fig2.text(0.51, 0.5, 'ACES Analytics Team', fontsize = 16, color = '#2ea7db', ha = 'center', va = 'top',
          alpha = 0.2)

plt.show()

"""
## 11. Format and save worksheets 
"""
## 11.1  Format worksheet - inventory ageing records
# Open blank work book
import xlwings as xw
wb = xw.Book()

# Define name of worksheet
sheet = wb.sheets["Sheet1"]
sheet.name = "Ageing Report"

# Hide gridlines for sheet
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

# Assign values of inventory records to worksheet
sheet.range("A1").options(index=False).value = ageing_report_ot

# define range of whole worksheet
data_rng = sheet.range("A1").expand('table')

# define height and width of each cell
data_rng.row_height = 23
data_rng.column_width = 13

# Format last row, and the previous row of last row
pr_len = str(len(ageing_report_ot) + 1)
last_row = "A"+ pr_len
last_row_e = "J"+ pr_len

pr_len_lastminus1 = str(len(ageing_report_ot))
lastminus1_row_e ="J"+ pr_len_lastminus1

bottom_rng = sheet.range(last_row , last_row_e)
for bi in range(7,11):
    bottom_rng.api.Borders(bi).Weight = 2
    bottom_rng.api.Borders(bi).Color = 0x70ad47
bottom_rng.api.Font.Name = 'Verdana'
bottom_rng.api.Font.Size = 9
bottom_rng.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng.api.Font.Color = 0x000000 #000000 #0xffffff
bottom_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

# Format border of cells
border_rng = sheet.range("A1", lastminus1_row_e)
for bi in range(7,13):
    border_rng.api.Borders(bi).Weight = 2
    border_rng.api.Borders(bi).Color = 0x70ad47
data_rng.api.Font.Name = 'Verdana'
data_rng.api.Font.Size = 9

# Format all range of the worksheet
data_rng.api.Font.Name = 'Verdana'
data_rng.api.Font.Size = 8
data_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
data_rng.api.WrapText = True

# Format header
header_rng = sheet.range("A1").expand('right')
header_rng.color = ('#70ad47') #cca989 #2da8dc
header_rng.api.Font.Color = 0xffffff
header_rng.api.Font.Bold = True
header_rng.api.Font.Size = 9
header_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

# Format first column
id_column_rng = sheet.range("A2").expand('down')
id_column_rng.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
id_column_rng.api.Font.Color = 0x000000 #000000 #0xffffff
id_column_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
id_column_rng.column_width = 10

# Format B column , Day
b_rng = sheet.range("B2").expand('down')
b_rng.column_width = 4

# Format C column , Profit Center
c_rng = sheet.range("C2").expand('down')
c_rng.column_width = 8
#c_debit_rng.number_format = "#,###.00"

# Format D column , Material Group
d_rng = sheet.range("D2").expand('down')
d_rng.column_width = 10
#d_credit_rng.number_format = "#,###.00"

# Format E column , Mtl Code
e_rng = sheet.range("E2").expand('down')
e_rng.column_width = 10
e_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

# Format F column , Mtl Text
f_rng = sheet.range("F2").expand('down')
f_rng.column_width = 16
#f_unsettled_rng.number_format = "#,###.00"

# Format G column , Volume\n (kg)
g_rng = sheet.range("G2").expand('down')
g_rng.number_format = "#,###"
g_rng.column_width = 10

# Format H column , Inventory Cost\n ($)
h_rng = sheet.range("H2").expand('down')
h_rng.number_format = "#,###"
h_rng.column_width = 15

# Format I column , Ageing Days
i_rng = sheet.range("I2").expand('down')
i_rng.column_width = 10

# Format J column , Ageing Buckets
j_rng = sheet.range("J2").expand('down')
j_rng.column_width = 10
j_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

# Add title
# Get length of rows
rowl = "{:,d}".format(len(ageing_report_ot))

# Get time
from time import strftime, localtime
time_local = strftime ("%A, %d %b %Y, %H:%M")

# Insert three rows above row 1
sheet.range("1:3").insert('down')
sheet.range("A1").value = "Inventory Ageing Report " + "As At " + closing_date_str
sheet.range("G1").value = "The last update time :  " + time_local  +"."
sheet.range("G2").value = "Updated by:  ACES Analytics Team."
sheet.range("A2").value = "Length of Rows: " + rowl

## 11.2  Format worksheet - Analysis
# Add new sheet
sheet2 = wb.sheets.add("Analysis")

# Hide gridlines for sheet2
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

# Hide gridlines
app = xw.apps.active
#app.api.ActiveWindow.DisplayGridlines = True
app.api.ActiveWindow.DisplayGridlines = False

## Format inventory analysis by ageing - start
sheet2.range("A1").options(index=True).value = ageing_sum_ot
sheet2.range("A1").value = "Ageing Bucket"

# Format common features such as font,font size
all_rng = sheet2.range("A1").expand('table')
all_rng.row_height = 23
all_rng.column_width = 10
all_rng.api.Font.Name = 'Verdana'
all_rng.api.Font.Size = 9
all_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
all_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

# Format boader
for bi in range(7,13):
    all_rng.api.Borders(bi).Weight = 2
    all_rng.api.Borders(bi).Color = 0x70ad47

# Format first column
col_1st = sheet2.range("A1").expand('down')
col_1st.column_width = 18

# Format header
header_rng = sheet2.range("A1").expand('right')
header_rng.color = ('#70ad47') #cca989 #2da8dc
header_rng.api.Font.Color = 0xffffff
header_rng.api.Font.Bold = True
header_rng.api.Font.Size = 9
header_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
header_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

# Format data area
data_rng = sheet2.range("B2").expand('table')
data_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
data_rng.number_format = "#,###"

# Insert one row to add title
sheet2.range("1:1").insert('down')
sheet2.range("A1").value = "Inventory Balance Analysis by Ageing"
title_rng = sheet2.range("A1")
title_rng.api.Font.Name = 'Verdana'
title_rng.api.Font.Size = 10
title_rng.api.Font.Color = 0x70ad47
title_rng.api.Font.Bold = True

## Format inventory analysis by ageing to worksheet - end

## Format inventory analysis by profit center - start
# Locate position for assign value
row_len = len(ageing_sum_ot) + 4
row_lenPlus1 =  row_len + 1
position = "A"+ str(row_len)
positionPLus1 = "A"+ str(row_lenPlus1)

#  Assign values to worksheet
sheet2.range(positionPLus1).options(index=True).value = prft_sum_ot
sheet2.range(positionPLus1).value = "Profit Center"

# Format common features
all_rng2 = sheet2.range(positionPLus1).expand('table')
all_rng2.row_height = 23
all_rng2.column_width = 10
all_rng2.api.Font.Name = 'Verdana'
all_rng2.api.Font.Size = 9
all_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
all_rng2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

# Format boader
for bi in range(7,13):
    all_rng2.api.Borders(bi).Weight = 2
    all_rng2.api.Borders(bi).Color = 0x70ad47

# Format header
header_rng2 = sheet2.range(positionPLus1).expand('right')
header_rng2.color = ('#70ad47') #cca989 #2da8dc
header_rng2.api.Font.Color = 0xffffff
header_rng2.api.Font.Bold = True
header_rng2.api.Font.Size = 9

# Format first column
col_1st = sheet2.range("A1").expand('down')
col_1st.column_width = 18

# Format data area
position_data = "B" + str(len(ageing_sum_ot) + 6)
data_rng = sheet2.range(position_data).expand('table')
data_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
data_rng.number_format = "#,###"

# Insert one row to add title
sheet2.range(position).value = "Inventory Balance Analysis by Profit Center"
title_rng = sheet2.range(position)
title_rng.api.Font.Name = 'Verdana'
title_rng.api.Font.Size = 10
title_rng.api.Font.Color = 0x70ad47
title_rng.api.Font.Bold = True

## Format inventory analysis by profit center - end

## Format inventory analysis by profit center and ageing in cost - start
# Locate position for assign value
row_len = row_len + len(prft_sum_ot) + 3
row_lenPlus1 = row_len +1
position = "A"+ str(row_len)
positionPLus1 = "A"+ str(row_lenPlus1)

#  Assign values to worksheet
sheet2.range(position).value = "Inventory Balance Analysis by Ageing and Profit Center in Cost ($)"
title_rng = sheet2.range(position)
title_rng.api.Font.Name = 'Verdana'
title_rng.api.Font.Size = 10
title_rng.api.Font.Color = 0x70ad47
title_rng.api.Font.Bold = True

sheet2.range(positionPLus1).options(index=True).value = ageing_prft_cost_ot

# Delete the first row of the worksheet
row_1st = sheet2.range(positionPLus1).expand('right')
row_1st.delete(shift='up')

# Format common features
all_rng2 = sheet2.range(positionPLus1).expand('table')
all_rng2.row_height = 23
all_rng2.column_width = 10
all_rng2.api.Font.Name = 'Verdana'
all_rng2.api.Font.Size = 9
all_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
all_rng2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

# Format boader
for bi in range(7,13):
    all_rng2.api.Borders(bi).Weight = 2
    all_rng2.api.Borders(bi).Color = 0x70ad47

# Format header
header_rng2 = sheet2.range(positionPLus1).expand('right')
header_rng2.color = ('#70ad47') #cca989 #2da8dc
header_rng2.api.Font.Color = 0xffffff
header_rng2.api.Font.Bold = True
header_rng2.api.Font.Size = 9

# Format first column
col_1st = sheet2.range("A1").expand('down')
col_1st.column_width = 18

# Format data area
position_data = "B" + str(row_lenPlus1 +1)
data_rng = sheet2.range(position_data).expand('table')
data_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
data_rng.number_format = "#,###"

# Format last row
pr_len2 = str(row_lenPlus1 + len(ageing_prft_cost_ot))
last_row2 = "A"+ pr_len2
bottom_rng2 = sheet2.range(last_row2).expand('right')

bottom_rng2.api.Font.Name = 'Verdana'
bottom_rng2.api.Font.Size = 9
bottom_rng2.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng2.api.Font.Color = 0x000000 #000000 #0xffffff

# Add value and format the cell
sheet2.range(positionPLus1).value = '              Profit Center\n Ageing'
A1_rng = sheet2.range(positionPLus1)
for bi in range(5, 6):
        A1_rng.api.Borders(bi).Weight = 2
        A1_rng.api.Borders(bi).Color = 0xffffff
## Format inventory analysis by profit center and ageing in cost - end

## Format inventory analysis by profit center and ageing in volume - start
# Locate position for assign value
row_len = row_len + len(ageing_prft_cost_ot) + 3
row_lenPlus1 = row_len +1
position = "A"+ str(row_len)
positionPLus1 = "A"+ str(row_lenPlus1)

#  Assign values to worksheet
sheet2.range(position).value = "Inventory Balance Analysis by Ageing and Profit Center in volume (kg)"
title_rng = sheet2.range(position)
title_rng.api.Font.Name = 'Verdana'
title_rng.api.Font.Size = 10
title_rng.api.Font.Color = 0x70ad47
title_rng.api.Font.Bold = True

sheet2.range(positionPLus1).options(index=True).value = ageing_prft_volume_ot

# Delete the first row of the worksheet
row_1st = sheet2.range(positionPLus1).expand('right')
row_1st.delete(shift='up')

# Format common features
all_rng2 = sheet2.range(positionPLus1).expand('table')
all_rng2.row_height = 23
all_rng2.column_width = 10
all_rng2.api.Font.Name = 'Verdana'
all_rng2.api.Font.Size = 9
all_rng2.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
all_rng2.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

# Format boader
for bi in range(7,13):
    all_rng2.api.Borders(bi).Weight = 2
    all_rng2.api.Borders(bi).Color = 0x70ad47

# Format header
header_rng2 = sheet2.range(positionPLus1).expand('right')
header_rng2.color = ('#70ad47') #cca989 #2da8dc
header_rng2.api.Font.Color = 0xffffff
header_rng2.api.Font.Bold = True
header_rng2.api.Font.Size = 9


# Format first column
col_1st = sheet2.range("A1").expand('down')
col_1st.column_width = 18

# Format data area
position_data = "B" + str(row_lenPlus1 +1)
data_rng = sheet2.range(position_data).expand('table')
data_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
data_rng.number_format = "#,###"

# Format last row
pr_len2 = str(row_lenPlus1 + len(ageing_prft_cost_ot))
last_row2 = "A"+ pr_len2
bottom_rng2 = sheet2.range(last_row2).expand('right')

bottom_rng2.api.Font.Name = 'Verdana'
bottom_rng2.api.Font.Size = 9
bottom_rng2.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
bottom_rng2.api.Font.Color = 0x000000 #000000 #0xffffff

# Add value and format the cell
sheet2.range(positionPLus1).value = '              Profit Center\n Ageing'
A1_rng = sheet2.range(positionPLus1)
for bi in range(5, 6):
        A1_rng.api.Borders(bi).Weight = 2
        A1_rng.api.Borders(bi).Color = 0xffffff
## Format inventory analysis by profit center and ageing in cost - end

# Format last column
# Get length of columns
len_col = len(ageing_prft_cost_ot.columns) + 1

# Convert column number to column name
def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

name_last_col = excel_column_name(len_col)
print(name_last_col)

# Format last column
row_lenPlus2 = str(row_lenPlus1 + 1)
Last_col_2ndcell = name_last_col + row_lenPlus2
last_col_rng = sheet2.range(Last_col_2ndcell).expand('down')
last_col_rng .color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
last_col_rng .api.Font.Color = 0x000000 #000000 #0xffffff

# Format first column
first_col_2ndcell = "A" + row_lenPlus2
first_col_rng = sheet2.range(first_col_2ndcell ).expand(('down'))
first_col_rng.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
first_col_rng.api.Font.Color = 0x000000 #000000 #0xffffff
first_col_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
first_col_rng.column_width = 20

# Format data area
first_cell_data_rng = "B" + row_lenPlus2
## Format inventory analysis by ageing and profit center - end
data_rng = sheet2.range(first_cell_data_rng).expand('table')
data_rng.number_format = "#,###"
data_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
## Format inventory analysis by profit center and ageing in volumn - end

# Get time
from time import strftime, localtime
time_local = strftime ("%A, %d %b %Y, %H:%M")

# Add plots
sheet2.range("1:28").insert('down')

sheet2.pictures.add(fig2, name='ACES_Plot2', update=True,
                     left=sheet2.range('A1').left, top=sheet2.range('A1').top)

sheet2.pictures.add(fig1, name='ACES_Plot1', update=True,
                     left=sheet2.range('G1').left, top=sheet2.range('G1').top)

# Add Title

sheet2.range("1:4").insert('down')
sheet2.range("A1").value = "Inventory Ageing Analysis" + "  " + "As At" + "  " + closing_date_str
sheet2.range("A2").value = "The last update time :  " + time_local  +"."
sheet2.range("A3").value = "Updated by:  ACES Analytics Team."

# Calculate running time to be shown in output file
end = timer()
running_time = "{:,.2f}".format(end - start)
sheet2.range("A4").value = "Running time:  " +  running_time + "  s"

"""
## 12. Export results to excel
"""
wb.save(output_file)
wb.close()
app.quit()

# Print the end for this script
print("The run of script is completed successfully.")
time_local_end = strftime ("%A, %d %b %Y, %H:%M")

# Print running time
running_time2 = "{:,.2f}".format(end - start)
print (running_time2 )

