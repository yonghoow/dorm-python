import pandas as pd
import xlsxwriter
import datetime

#read the excel file
df = pd.read_excel('sample2.xlsx', skiprows=1)

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
writer = pd.ExcelWriter('formatting.xlsx', engine='xlsxwriter')

# Write the dataframe data to xlsxwriter
new_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# add a border format to use to highlight cells.
border = workbook.add_format({'border': 1})

# set column width
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 20)

# Write some data to the worksheet.
worksheet.write('A5', None, border)
worksheet.write('B3', None, border)

#close the Pandas Excel writer and output the Excel file
workbook.close()