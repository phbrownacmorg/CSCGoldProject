from email.message import EmailMessage
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
from_addr = 'peter.brown@converse.edu'

def verify_constants() -> bool:
    """Verify that all the constant paths exist.  If any doesn't, raise AssertionError."""
    assert dropbox_dir.is_dir(), 'Dropbox directory not found'
    assert input_folder.is_dir(), 'Input folder not found'
    assert processed_folder.is_dir(), 'Processed folder not found'
    assert csv_folder.is_dir(), "CSV folder not found"
    assert instructors_file.is_file(), 'instructors.csv not found'
    assert coursenames_file.is_file(), 'coursenames.csv not found'
    return True

def normalize_filename(f: Path) -> Path:
    """Takes a Path F and normalizes the filename to remove unwanted characters, notably spaces.
        Renames the file F to the normalized name."""
    badchars = ' ' # More characters could be added if need be
    newname = f.name
    for c in badchars:
        newname = newname.replace(c, '')
    return f.replace(f.parent.joinpath(newname))

def list_input_files() -> list[Path]:
    """Return the list of .XLSX files in input_folder, sorted alphabetically."""
    files: list[Path] = list(input_folder.glob('*.xlsx'))
    for i in range(len(files)):
        files[i] = normalize_filename(files[i])
    return sorted(files) # To give a stable ordering

def make_frame(csv_file: Path, idx_label: str) -> pd.DataFrame:
    frame = pd.read_csv(csv_file)
    frame = frame.set_index(idx_label)
    return frame

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
    d['coursename'] = cnm_df.loc[d['display_num']].coursename       # type: ignore
    d['term'] = infile_parts[2]
    d['full_num'] = d['prefix'] + d['barenum'] + '.' + d['section'] + '-' + d['term']
    d['students_csv'] = str(csv_folder.joinpath(d['full_num'] + '.csv'))
    d['instructor'] = cast(pd.DataFrame, i_df.loc[infile_parts[3]])
    return d

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

    # Should filter data here
    # 1. Fix capitalization of column names? Might could use columns and then reindex
    # 2. Ensure names begin with a capital letter (*not* title-case!)
    # 3. Dump anything in "email address" that doesn't end in "@converse.edu"
    # 4. Dump anything in "School email" that *does* end in "@converse.edu"

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
    coursenum = course_dict['display_num'] + '.' + course_dict['section']
    lastname = course_dict['instructor'].at['lastname']                                                     # type: ignore

    msg = EmailMessage()
    msg['Subject'] = coursenum
    print('Subject:', msg['Subject'])
    msg['From'] = from_addr
    msg['To'] = course_dict['students'].loc[:, ['email address']].to_dict(orient='list')['email address']   # type: ignore
    print('To:', msg['To'])
    msg['CC'] = course_dict['instructor'].at['email']                                                       # type: ignore
    print('CC:', msg['CC'])
    msg['BCC'] = course_dict['students'].loc[:, ['School EMAIL']].to_dict(orient='list')['School EMAIL']    # type: ignore
    print('BCC:', msg['BCC'])

    msg.set_content(f"""\
Welcome!  You are registered for {coursenum}, {course_dict['coursename']}. Your professor will be Dr. {course_dict['instructor'].at['firstname']} {lastname} ({msg['CC']}).

Your course is being taught through Converse's Canvas. The simplest way to get in to Converse's Canvas is to go to https://converse.instructure.com and log in there with your Converse email and password.

If you've taken a course with Converse before, your Converse email address and password will normally be unchanged from what they were. If you have a new Converse account, you should have gotten your Converse email address and password in a separate email from Campus Technology.

Once logged in, you should be taken to your Canvas dashboard. On that dashboard, you should see a tile with the name of your course.

If you don't see the tile before the course starts, that is not (yet) a problem. Our Canvas courses are created unpublished, which means they're hidden from students.

Please remember that you can always email me (peter.brown@converse.edu) for Canvas questions. Dr. {lastname} is a better source for all other questions.

Peace,
—Peter Brown

Peter H. Brown, Ph.D.
Asst. Professor of Computer Science
Director of Distance Education
Converse University
""")
    print(msg.get_content())
   
def write_students_csv(course_dict: dict[str, str | pd.DataFrame]) -> None:
    """Takes a dict COURSE_DICT and writes COURSE_DICT['students'] to the
        appropriate CSV file."""
    course_dict['students'].to_csv(course_dict['students_csv'], quoting=csv.QUOTE_STRINGS) # type: ignore

def move_to_processed(course_dict: dict[str, str | pd.DataFrame]) -> None:
    """Takes a Path INFILE and moves it to processed_folder."""
    infile = Path(cast(str, course_dict['infile']))                    
    infile.replace(processed_folder.joinpath(infile.name))
    
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
        move_to_processed(course_dict)
    return 0

if __name__ == '__main__':
    import sys
    sys.exit(main(sys.argv))

# TODO:
# 1. Filter column headings
# 2. Remove non-Converse email addresses from To:
# 3. Remove Converse email addresses from BCC:
# 4. Capitalize names
# 5. Correctly handle the two EDU courses
# 6. Actually send mail?