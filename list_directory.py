import glob
import os

#read latest file with format sample*.xlsx

def read_latest_file(pattern):
    for file_path in glob.glob(pattern):
        
        print('Found file pattern: ', file_path)
        latest_file = max(glob.glob(pattern), key=os.path.getctime)
        print('latest file: ', latest_file)

read_latest_file('sample*.xlsx')