import pandas as pd

df = pd.read_excel('sample2.xlsx', skiprows=1)

#sorting by column 'UNIT NO'
df.sort_values(by='UNIT NO', inplace=True)

print(df)

#remove duplicate 'Unit No' column and keep the first occurence
unit_no_list = list(dict.fromkeys(df['UNIT NO']))
print(unit_no_list)

empty_rows = 2

#create empty dataframe with empty rows
empty_data = {col: ['']*empty_rows for col in df.columns}
empty_df = pd.DataFrame(empty_data)

#filter the dataframe by unit no
for x in unit_no_list:
    filtered_df = df[df['UNIT NO'] == x]
    #insert 2 blank rows
    filtered_df = pd.concat([filtered_df, empty_df], ignore_index=True)
    print(filtered_df)

#save to new excel file
#df.to_excel('sample_sorted.xlsx', index=False)