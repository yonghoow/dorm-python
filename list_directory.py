import glob
import os

#read latest file with format Door*.xls

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


