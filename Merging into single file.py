import pandas as pd
import os
from pathlib import Path

# user input for the folder path
r = "r'"
path = Path(input('Choose a folder:'))

paths = os.path.join(r, path)

# change the current directory to the user-specified path
os.chdir(paths)

# return file names from the path to a list; index tuples from os.walk
NDC = []
for x in os.walk(os.getcwd()):
    NDC.extend(x[2])

# create a list with file names joined with the path
ndc_paths = []
for ndc in NDC:
    ndc_paths.append(os.path.join(str(path), ndc))

# load all files from the path into Python using pandas and assign each a dataframe number
dataframes = {}
for i, file in enumerate(ndc_paths):
    df_name = f'df{i+1}'
    dataframes[df_name] = pd.read_excel(file)

# extract the NDC number from each dataframe
ndc_name = []
for key, item in dataframes.items():
    #ndc.append(item.iloc[5,1])
    ndc_name.append(item.iloc[5,1])

# user input for the name of the new xlsx file
excel_file_name = input('What would you like to name the file? Add the ".xlsx" extension:')

# create the path for the new xlsx file
excel = os.path.join(path, excel_file_name)

# combine all dataframes into a single xlsx and create a new xlsx file
with pd.ExcelWriter(excel) as writer:
    for df_name, sheet_name in zip(dataframes.keys(), ndc_name):
        dataframes[df_name].to_excel(writer, sheet_name=sheet_name, index=False)
