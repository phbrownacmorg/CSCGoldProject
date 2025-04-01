#import pandas as pd 

#dataframe1 = pd.read_excel('EDU591.Y2_Students_anonymized.xlsx')

#print(dataframe1)

import pandas as pd
import os
import glob

excel_files = glob.glob('*.xlsx') + glob.glob('*.xls')

if not excel_files:
    print("No Excel files found in the current directory.")
else:
    print(f"Found {len(excel_files)} Excel file(s):")

    for excel_file in excel_files:
        print(f"Processing: {excel_file}")
   
        df = pd.read_excel(excel_file)
        
       
        csv_file = os.path.splitext(excel_file)[0] + '.csv'
      
        df.to_csv(csv_file, index=False)
        
        print(f"Successfully converted to: {csv_file}")
        print(f"Preview of the data:\n{df.head()}\n")
