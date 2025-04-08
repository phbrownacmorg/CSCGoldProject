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
        
        base_name = os.path.splitext(excel_file)[0]
        parts = base_name.split('_')
        
        if parts and ',' in parts[-1]:
            prof_last_name = parts[-1].split(',')[0]
        else:
            prof_last_name = 'Unknown'

        print(f"Professor Last Name Detected: {prof_last_name}")

        df = pd.read_excel(excel_file)

        csv_file = base_name + '.csv'

        df.to_csv(csv_file, index=False)

        print(f"Successfully converted to: {csv_file}")
        print(f"Preview of the data:\n{df.head()}\n")
