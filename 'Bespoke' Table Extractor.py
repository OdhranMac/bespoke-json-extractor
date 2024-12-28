"""
This script converts a json file into an excel file.
Specify json file in rootDir, output directory/filename in outputDir.
"""

import json
import pprint as pp
import pandas as pd
import numpy as np

array = np.chararray(shape = (50, 20), itemsize = 1000, unicode = True) 
rootDir = r''
outputDir = r''

# get table function
def get_table(JSON_data):
    """
    Converts json structure into numpy array.
    
    args:
        JSON_data (json)
        
    returns:
        array (numpy array)
    """

    pages = JSON_data['tables']
    correctTable = False # let's program know we have correct table

    for page in pages:
        print('page #: ' + str(page))
        tables = pages[page]
        
        for table in tables:
            cells = table['cells']

            for cell in cells:
                
                if (cell['row'] == 0 and cell['column'] == 0 and cell['text'] == 'Company' or correctTable == True):

                    # guarantee table is not other one on page:
                    if (cell['row'] == 0 and cell['column'] == 1 and cell['text'] == 'Direct-Man HRS'):
                        correctTable = False
                        array[0, 0] = ''
                        continue
                    else:
                        pass

                    correctTable = True            
                    
                    # get data from this table then break
                    value = cell['text']
                    array[cell['row'], cell['column']] = value    
                else:
                    continue

                if ((correctTable == True) and (cell['row'] == table['rows'] - 1) and (cell['column'] == table['columns'] - 1)):
                    return array

def main():
    JSON_file = open(rootDir) 
    JSON_data = json.load(JSON_file)  

    array = get_table(JSON_data)
    print(array)

    df = pd.DataFrame(data = array, dtype = str, index = None, columns = None)
    df.to_excel(outputDir, header = False, index = False)

if __name__ == "__main__":
    main()