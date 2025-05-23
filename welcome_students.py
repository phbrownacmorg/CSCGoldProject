from pathlib import Path
from typing import cast
import csv
import pandas as pd

# Constants
dropbox_dir = Path.home().joinpath('Dropbox')
input_folder = dropbox_dir.joinpath('DEd', 'students_in')
code_folder = Path(__file__).parent
processed_folder = code_folder.joinpath('students_processed')
csv_folder = code_folder.joinpath('students_csv')
instructors_file = csv_folder.joinpath('instructors.csv')
coursenames_file = csv_folder.joinpath('coursenames.csv')

def verify_constants() -> bool:
    """Verify that all the constant paths exist.  If any doesn't, raise AssertionError."""
    assert dropbox_dir.is_dir(), 'Dropbox directory not found'
    assert input_folder.is_dir(), 'Input folder not found'
    assert processed_folder.is_dir(), 'Processed folder not found'
    assert csv_folder.is_dir(), "CSV folder not found"
    assert instructors_file.is_file(), 'instructors.csv not found'
    assert coursenames_file.is_file(), 'coursenames.csv not found'
    return True

def list_input_files() -> list[Path]:
    """Return the list of .XLSX files in input_folder, sorted alphabetically."""
    files: list[Path] = list(input_folder.glob('*.xlsx'))
    return sorted(files) # To give a stable ordering

def make_frame(csv_file: Path, idx_label: str) -> pd.DataFrame:
    frame = pd.read_csv(csv_file)
    frame = frame.set_index(idx_label)
    return frame

def make_coursename_frame() -> pd.DataFrame:
    frame = pd.read_csv(instructors_file)
    frame = frame.set_index('filename')
    return frame

# def get_coursename(display_number: str) -> str:
#     result = ''
#     with open(csv_folder.joinpath('coursenames.csv')) as f:
#         reader = DictReader(f)
#         for row in reader:
#             if display_number == row['course']:
#                 result = row['coursename']
#                 break
#     # Post: the course was found
#     assert result != '', f"Course {display_number} not found in coursenames.csv"
#     return result

# def get_instructor(instructor_code: str, instructor_frame: pd.DataFrame) -> pd.DataFrame:
#     # result: dict[str, str] = {}
#     # instructor_frame = pd.read_csv(instructors_file)
#     instructor_frame = instructor_frame.set_index('filename')
#     print(instructor_frame)
#     print()
#     # print(instructor_frame.axes)
#     result = instructor_frame.loc[instructor_code]
#     print(result)
#     print()
#     return result



    # with open(csv_folder.joinpath('instructors.csv')) as f:
    #     reader = DictReader(f)
    #     for row in reader:
    #         if instructor_code == row['filename']:
    #             result = row
    #             break
    # # Post: the instructor was found
    # assert result != {}, f"Instructor '{instructor_code}' not found in instructors.csv"
    # return result

def init_course_dict(infile: Path, i_df: pd.DataFrame, cnm_df: pd.DataFrame) -> dict[str, str | pd.DataFrame]:
    """Initialize and return a dictionary with those properties of the course
        that can be inferred from the INFILE."""
    d: dict[str, str | pd.DataFrame] = {}
    d['infile'] = str(infile)
    infile_parts = infile.stem.split('_')
    assert len(infile_parts) == 4
    d['prefix'] = infile_parts[0][:3]
    dot = infile_parts[0].find('.')
    if dot < 0:
        dot = len(infile_parts[0])
    assert dot > 0
    d['barenum'] = infile_parts[0][3:dot]
    d['section'] = infile_parts[0][dot+1:]
    d['display_num'] = d['prefix'] + ' ' + d['barenum']
    d['coursename'] = cast(str, cnm_df.loc[d['display_num']].coursename)
    d['term'] = infile_parts[2]
    d['full_num'] = d['prefix'] + d['barenum'] + '.' + d['section'] + '-' + d['term']
    d['students_csv'] = str(csv_folder.joinpath(d['full_num'] + '.csv'))

    d['instructor'] = cast(pd.DataFrame, i_df.loc[infile_parts[3]])


    return d

# def corresponding_csv(infile: Path) -> Path:
#     """Takes a path INFILE and returns the filename of the corresponding CSV
#         file in csv_folder, whether or not that file exists."""
#     # Really should parse INFILE
#     csv_path = csv_folder.joinpath(infile.name).with_suffix('.csv')
#     return csv_path

def read_input(infile: Path, csvfile: str) -> pd.DataFrame:
    """Takes a path INFILE and reads it into a pandas DataFrame.  If a 
        corresponding CSV file exists in csv_folder, use the information from
        there to update the DataFrame.  Return the DataFrame."""
    # Should really look at the file and figure out the value for header
    # Remember that header is 0-indexed, and so one less than the row number
    frame = pd.read_excel(infile, header=2)
    # Get rid of the empty column A, if present
    frame = frame.dropna(axis='columns',how='all')
    frame = frame.set_index('Jenzabar ID')

    #csvfile = corresponding_csv(infile)
    if Path(csvfile).is_file():
        # Read that, and update frame with the data
        frame2 = pd.read_csv(csvfile)
        frame2 = frame2.set_index('Jenzabar ID')
        frame = frame.combine_first(frame2)

    return frame

def send_emails(course_dict: dict[str, str | pd.DataFrame]) -> None:
    """Takes a dict COURSE_DICT and uses it to send welcome emails to the
        students listed in COURSE_DICT['students'] who have not already 
        received emails.  COURSE_DICT['students'] is updated with the 
        information that the new emails have been sent."""
    
def write_students_csv(course_dict: dict[str, str | pd.DataFrame]) -> None:
    """Takes a dict COURSE_DICT and writes COURSE_DICT['students'] to the
        appropriate CSV file."""
    course_dict['students'].to_csv(course_dict['students_csv'], quoting=csv.QUOTE_STRINGS) # type: ignore

def move_to_processed(infile: Path) -> None:
    """Takes a Path INFILE and moves it to processed_folder."""

def main(args: list[str]) -> int:
    verify_constants()
    input_files = list_input_files()
    i_df = make_frame(instructors_file, 'filename')
    cnm_df = make_frame(coursenames_file, 'course')
    for f in input_files:
        course_dict: dict[str, str | pd.DataFrame] = init_course_dict(f, i_df, cnm_df)
        course_dict['students'] = read_input(f, cast(str, course_dict['students_csv']))
        send_emails(course_dict)
        write_students_csv(course_dict)
        move_to_processed(f)

    return 0

if __name__ == '__main__':
    import sys
    sys.exit(main(sys.argv))