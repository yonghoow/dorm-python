import pandas as pd
import xlsxwriter
import datetime
import glob
import os

#read latest file with format door.xls

def read_latest_file(pattern):
    for file_path in glob.glob(pattern):
        
        print('Found file pattern: ', file_path)
        latest_file = max(glob.glob(pattern), key=os.path.getctime)
        print('latest file: ', latest_file)

    return latest_file    

#get the filename
pattern = '../Door*.xls'
file = read_latest_file(pattern)
print('File: ', file)

#read the excel file
df = pd.read_excel(file, skiprows=1)

#remove all rows with '--' in 'UNIT NO' column
df = df.drop(df[df['UNIT NO'] == '--'].index)


#sorting by column 'UNIT NO'
df.sort_values(by='UNIT NO', inplace=True)

#remove duplicate 'Unit No' column and keep the first occurence
unit_no_list = list(dict.fromkeys(df['UNIT NO']))

#create empty dataframe with 1 empty rows
empty_data = {col: ['']*1 for col in df.columns}
empty_df = pd.DataFrame(empty_data)

#create new dataframe
new_df = pd.DataFrame(columns=df.columns)

#create headings for dataframe
new_heading = pd.DataFrame({'UNIT NO': ['UNIT NO'], 'NAME OF SPONSOR COMPANY': ['NAME OF SPONSOR COMPANY'], 'NAME OF WORKERS': ['NAME OF WORKERS'], 'FIN NUMBER': ['FIN NUMBER'], 'WP EXPIRY': ['WP EXIPRY'], 'CHECK IN DATE': ['CHECK IN DATE'], 'BED NO': ['BED NO']})

#create footer for the dataframe
current_day = datetime.datetime.now().strftime('1/%m/%Y')
footer = pd.DataFrame({'UNIT NO': [''], 'NAME OF SPONSOR COMPANY': [''], 'NAME OF WORKERS': [''], 'FIN NUMBER': [''], 'WP EXPIRY': [''], 'CHECK IN DATE': [''], 'BED NO': [current_day]})

# count iteration using j
j = 0

#filter the dataframe by unit no
for x in unit_no_list:
    filtered_df = df[df['UNIT NO'] == x]
    #count to 12 rows
    rows = len(filtered_df)
    if rows < 12:
        for i in range(12-rows):
            filtered_df = pd.concat([filtered_df, pd.DataFrame([{'UNIT NO': x, 'NAME OF SPONSOR COMPANY': '', 'NAME OF WORKERS': '', 'FIN NUMBER': '', 'WP EXPIRY': '', 'CHECK IN DATE': '', 'BED NO': ''}])], ignore_index=True)
    #insert footer
    filtered_df = pd.concat([filtered_df, footer], ignore_index=True)
    #insert empty rows using while loop

    if j % 2 != 0:
        filtered_df = pd.concat([filtered_df, empty_df], ignore_index=True)
    else :
        filtered_df = pd.concat([filtered_df, empty_df, empty_df, empty_df, empty_df], ignore_index=True)
    #increment i    
    j += 1

    #append to new dataframe
    new_df = pd.concat([new_df, new_heading, filtered_df], ignore_index=True)

#create a writer object using xlsxwriter as the engine
writer = pd.ExcelWriter('../door_formatted.xlsx', engine='xlsxwriter')

# Write the dataframe data to xlsxwriter
new_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)


# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# add a border format to use to highlight cells.
border = workbook.add_format({'border': 1})

#print the number of rows in the dataframe
rows = len(new_df)
print('no. of rows in dataframe: ', rows)

#conditional formatting
# Add a format. Border is needed for the conditional formatting to work

#create a while loop to iterate through the rows
A_value = 1 #initial value
G_value = 13 #initial value
k = 0 #count iteration, check if k is even or odd

while G_value < rows:
    if k % 2 == 0:
        firstCol = 'A'+ str(A_value) + ':G' + str(G_value)
        worksheet.conditional_format(firstCol, {'type' : 'text', 'criteria' : 'not containing', 'value' : '/', 'format' : border})
        A_value += 18
        G_value += 18
        k += 1
        
    else :
        firstCol = 'A'+ str(A_value) + ':G' + str(G_value)
        worksheet.conditional_format(firstCol, {'type' : 'text', 'criteria' : 'not containing', 'value' : '/', 'format' : border})
        A_value += 15
        G_value += 15
        k += 1


# set column width
worksheet.set_column('A:A', 7.7)
worksheet.set_column('B:B', 55.7)
worksheet.set_column('C:C', 32.6)
worksheet.set_column('D:D', 10.7)
worksheet.set_column('E:E', 9.6)
#wrap text for column F
wrap_format = workbook.add_format({'text_wrap': True})
worksheet.set_column('F:F', 9.4, wrap_format)
worksheet.set_column('G:G', 8.1)

#set row height
worksheet.set_default_row(33)



#close the Pandas Excel writer and output the Excel file
workbook.close()

