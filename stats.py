import pandas as pd
import numpy as np 
import xlsxwriter
import sys
import datetime

# File name of created workbook includes current date
now = datetime.datetime.now()
Excel_name = "ProposedMitigations_"+now.strftime("%Y%b%d")+".xlsx"
workbook = xlsxwriter.Workbook(Excel_name)
worksheet = workbook.add_worksheet('FinalReport')

# read input files, csv is Veracode data export; Excel is prior spreadsheet to carry forward status notes
# note that csv file is generated from Veracode export by ScanMitigationcsv.py
df = pd.read_csv('ProposedMitigations_2019May28.csv', encoding = "cp1252")
prpt = pd.ExcelFile('ProposedMitigations_2019May23.xlsx')
old_rpt = prpt.parse('FinalReport')
#old_rpt = old_rpt.drop(['0 - Informational',  '1 - Very Low','2 - Low','3 - Medium', '4 - High','5 - Very High','Grand_Total' ], axis=1)
old_rpt = old_rpt.drop(['0 - Informational', '1 - Very Low','2 - Low','3 - Medium', '4 - High','5 - Very High','Grand_Total' ], axis=1)

# create table of app-scan to Business
xref = df[['App_Name_and_Build','Business']]
xref = xref.drop_duplicates()
xref = xref.set_index('App_Name_and_Build')
#row = xref.loc['989 - UW-GAIN - 14 Mar 2019 Static (2)']
#print (row['Business'])
#exit()

# Count proposed mitigations by severity for each scan
xtab = pd.crosstab(index=df.App_Name_and_Build,columns=df.Severity_Label)

# Carry forward prior status notes, including analyst assigned to scan
#xtab_merged = pd.merge(xtab, old_rpt, on='App_Name_and_Build')
xtab_merged = pd.merge(xtab, old_rpt, on='App_Name_and_Build', how='left')
xtab_merged = xtab_merged.fillna('NA')

# Define Cell Formats
bold = workbook.add_format({'bold': True})
bold.set_bg_color('#D3DDEC')
header_format = workbook.add_format({'center_across': True,'bold': True, 'border': True})
header_format.set_bg_color('#D3DDEC')
cell_format_ctr = workbook.add_format({'center_across': True, 'border': True})
cell_format_lft = workbook.add_format({ 'border': True})
cell_format_dt = workbook.add_format({'center_across': True, 'border': True})
cell_format_dt.set_num_format('mm/dd/yy')

#define column widths
worksheet.set_column(0,0,65)
worksheet.set_column(1,7,15)
worksheet.set_column(8,8,18)
worksheet.set_column(9,10,12)
worksheet.set_column(11,11,65)

# write worksheet header line
worksheet.write(0,0,'App_Name_and_Build',bold)
worksheet.write(0,1,'0 - Informational',header_format)
worksheet.write(0,2,'1 - Very Low',header_format)
worksheet.write(0,3,'2 - Low',header_format)
worksheet.write(0,4,'3 - Medium',header_format)
worksheet.write(0,5,'4 - High',header_format)
worksheet.write(0,6,'5 - Very High',header_format)
worksheet.write(0,7,'Grand_Total',header_format)
worksheet.write(0,8,'Business',header_format)
worksheet.write(0,9,'Assignee',header_format)
worksheet.write(0,10,'Status_Date',header_format)
worksheet.write(0,11,'Status_Notes',bold)


	
# go through dataframe and output in Excel
row_nbr = 1
for index, row in xtab_merged.iterrows():
	worksheet.write(row_nbr,0,row['App_Name_and_Build'],cell_format_lft)
	try:
		worksheet.write(row_nbr,1,row['0 - Informational'], cell_format_ctr)
	except:
		row['0 - Informational']= 0
		worksheet.write(row_nbr,1,0, cell_format_ctr)
	try:
		worksheet.write(row_nbr,2,row['1 - Very Low'], cell_format_ctr)
	except:
		row['1 - Very Low']= 0
		worksheet.write(row_nbr,2,0, cell_format_ctr)	
	worksheet.write(row_nbr,3,row['2 - Low'], cell_format_ctr)
	worksheet.write(row_nbr,4,row['3 - Medium'], cell_format_ctr)
	worksheet.write(row_nbr,5,row['4 - High'], cell_format_ctr)
	try:
		worksheet.write(row_nbr,6,row['5 - Very High'], cell_format_ctr)
	except:
		row['5 - Very High']= 0
		worksheet.write(row_nbr,6,0, cell_format_ctr)	
	row_total = row['0 - Informational']+ row['1 - Very Low']  + row['2 - Low'] + row['3 - Medium'] + row['4 - High'] + row['5 - Very High']
	worksheet.write(row_nbr,7,row_total, cell_format_ctr)
	biz = xref.loc[row['App_Name_and_Build']]
	worksheet.write(row_nbr,8,biz['Business'], cell_format_ctr)
	worksheet.write(row_nbr,9,row['Assignee'], cell_format_ctr)
	worksheet.write(row_nbr,10,row['Status_Date'], cell_format_dt)
	worksheet.write(row_nbr,11,row['Status_Notes'], cell_format_lft)
	row_nbr += 1
workbook.close()

