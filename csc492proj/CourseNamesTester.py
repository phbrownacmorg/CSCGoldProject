import pandas as pd
import glob

courseNames_df = pd.read_csv("coursenames.csv")

excel_files = glob.glob('*.xlsx') + glob.glob('*.xls')

if not excel_files:
    print("No Excel files found in the current directory.")
else:
    print(f"Found {len(excel_files)} Excel file(s):")

    courseNames_df.iloc[:, 0] = courseNames_df.iloc[:, 0].astype(str)

    for excel_file in excel_files:
        print(f"\nProcessing: {excel_file}")

        temp = pd.read_excel(excel_file, header=None, nrows=5)

        courseCode = temp.iloc[0, 0]
        courseCode = str(courseCode)
        courseCode = courseCode[:7]  # Keep only the first 7 characters, which should be 3 letters, a space, and then 3 numbers.
        print(f"Extracted Course Code: {courseCode}")

        course_row = courseNames_df[courseNames_df.iloc[:, 0].str.strip() == courseCode]

        if not course_row.empty:
            courseName = course_row.iloc[0, 1]
            print(f"Course Name Found: {courseName}")
        else:
            courseName = "Unknown Course Name"
            print(f"Course name not found for code: {courseCode}")