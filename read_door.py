import pandas as pd
import xlsxwriter
import datetime

df = pd.read_excel('sample2.xlsx', skiprows=1)

#sorting by column 'UNIT NO'
df.sort_values(by='UNIT NO', inplace=True)

print(df)

#remove duplicate 'Unit No' column and keep the first occurence
unit_no_list = list(dict.fromkeys(df['UNIT NO']))
print(unit_no_list)


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
    
    #print filtered dataframe
    print(filtered_df)
    #append to new dataframe
    new_df = pd.concat([new_df, new_heading, filtered_df], ignore_index=True)


#print new dataframe
print(new_df)


#save to new excel file
new_df.to_excel('sample_sorted.xlsx', index=False, header=False)
