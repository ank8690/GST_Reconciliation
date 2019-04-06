"""
    This is a script to loop through all the excel files in the current working directory specifically related to GSTR2A.
	The next target is to unmerge the cells which are in the merged conditions and write the unmerged file to excel.
	Then delete the unnecessary rows and finally add these files into one excel final file.
    The Excel file used in this script can be found here:
        http://www.unicef.org/sowc2014/numbers/
"""
import pandas as pd
import numpy as np
import glob
import os
import xlrd
import xlwt
path=os.getcwd() # to get the cwd (current working directory) path
files = os.listdir(path)
files_xlsx = [f for f in files if f[-4:] == 'xlsx']
print(files_xlsx)
unmerged_files = []
for f in files_xlsx:
	print(f)
	if not os.path.exists(f):
		print("Could not find the excel file: " % f)
		continue

    # read merged cells for all sheets
    #book = xlrd.open_workbook(path, formatting_info=True)
	book = xlrd.open_workbook(f)
	writed_cells = []   # writed cell for merged cells
    # open excel file and write
	excel = xlwt.Workbook()
	for rd_sheet in book.sheets():
		if rd_sheet.name != 'B2B':
			continue
        # for each sheet
		print(rd_sheet.name)
		wt_sheet = excel.add_sheet(rd_sheet.name) # this is the sheet to be written in new excel file
		#writed_cells = []
		#print(rd_sheet.merged_cells)
		#size = len(rd_sheet.merged_cells)
		#print(size)
        # over write for merged cells
		for crange in rd_sheet.merged_cells[2:]:
            # for each merged_cell
			rlo, rhi, clo, chi = crange # for each merged cell, there is row range(low to high) as well as column range.
			cell_value = rd_sheet.cell(rlo, clo).value
			#print(cell_value)
			if cell_value != 'Invoice details' and cell_value != 'Tax Amount': 
				wt_sheet.write(rhi-1, chi-1, cell_value)
				writed_cells.append((rhi-1, chi-1))
				
			
        # write all un-merged cells
		
		for r in range(5, rd_sheet.nrows):
			for c in range(0, rd_sheet.ncols):
				if (r, c) in writed_cells:
					continue
				invoice_number = rd_sheet.cell(r, 2).value # This is to get the current row invoice number 
				prev_invoice_Number = rd_sheet.cell(r-1, 2).value + '-Total'
				if invoice_number == prev_invoice_Number or invoice_number == rd_sheet.cell(r-1, 2).value:
					continue
				cell_value = rd_sheet.cell(r, c).value
				#print(cell_value)
				#if cell_value == '':
				#	continue
				wt_sheet.write(r, c, cell_value)

	(origin_file, ext) = os.path.splitext(f)
	unmerge_excel_file = origin_file + '_unmerged' + ext
	excel.save(unmerge_excel_file)
	unmerged_files.append(unmerge_excel_file)


print(unmerged_files)
# Making a list of missing value types
missing_values = ["Unnamed: 0", "na", "--"]
new_excel_names = []
for name in unmerged_files:
	data = pd.read_excel(name, na_values = missing_values) 
	# read_excel and read_csv works the same way. They have the same attributes like .head(), .describe()
	#print(data.head())
	#print(data.describe())
	#print(data.shape)
	#empty = data.isnull()
	#df2 = data[data.isnull() == False]
	df2 = data.ix[1:]
	df2 = data.ix[2:]
	df2 = data.ix[3:]
	df2 = data.ix[4:]
	#df2.filter(df2.iloc[0])
	
	(origin_file, ext) = os.path.splitext(name)
	del_excel_file = origin_file + '_deleted' + ext
	df2.to_excel(del_excel_file, header=False, index=False)
	new_excel_names.append(del_excel_file)
	#print(df2.head())

#new_excel_names = ["05AAHFJ2458G1ZG_102018_R2A_unmerged_deleted.xlsx", "05AAHFJ2458G1ZG_112018_R2A_unmerged.xlsx"]
excels = [pd.ExcelFile(name) for name in new_excel_names]
# turn them into dataframes
frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

# delete the first row for all frames except the first
# i.e. remove the header row -- assumes it's the first
frames[1:] = [df[1:] for df in frames[1:]]

# concatenate them..
combined = pd.concat(frames)

# write it out
combined.to_excel("c.xlsx", header=False, index=False)
