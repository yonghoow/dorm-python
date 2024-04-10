import pandas as pd

df = pd.read_excel('sample.xlsx', skiprows=1)

#sorting by column 'UNIT NO'
df.sort_values(by='UNIT NO', inplace=True)

print(df)

#reset index
df.reset_index(drop=True, inplace=True)

print('After resetting index:')
print(df)

#save to new excel file
df.to_excel('sample_sorted.xlsx', index=False)