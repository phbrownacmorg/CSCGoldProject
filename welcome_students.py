from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from smtplib import SMTP
from typing import cast
import csv
import pandas as pd
import re

# Constants
dropbox_dir = Path.home().joinpath('Dropbox')
input_folder = dropbox_dir.joinpath('DEd', 'students_in')
code_folder = Path(__file__).parent
processed_folder = code_folder.joinpath('students_processed')
csv_folder = code_folder.joinpath('students_csv')
instructors_file = csv_folder.joinpath('instructors.csv')
coursenames_file = csv_folder.joinpath('coursenames.csv')
from_addr = 'peter.brown@converse.edu'
files_to_skip = ['CSC000.00_2025-06-01_2425-BS_Brown,P.']

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

def make_start_date(spec: str) -> str:
    try:
        result = date.fromisoformat(spec)
    except ValueError:
        result = date.today()
    return result.isoformat()

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
    d['start_date'] = make_start_date(infile_parts[1])
    return d

def cap_first(s: str) -> str:
    """Takes a string S and returns S with its first character capitalized and the rest left alone."""
    if len(s) > 0:
        s = s[0].title() + s[1:]
    return s

def read_input(infile: Path, csvfile: str) -> pd.DataFrame:
    """Takes a path INFILE and reads it into a pandas DataFrame.  If a 
        corresponding CSV file exists in csv_folder, use the information from
        there to update the DataFrame.  Return the DataFrame."""
    # Should really look at the file and figure out the value for header
    # Remember that header is 0-indexed, and so one less than the row number
    frame = pd.read_excel(infile, header=2, dtype='str')
    frame = frame.dropna(axis='columns',how='all') # Get rid of empty columns
    frame = frame.set_index('Jenzabar ID')
    frame['sent'] = None

    # Fix column names
    new_cols = {'FIRST NAME': 'First name', 'LAST NAME': 'Last name', 
               'School EMAIL': 'school email', 'Recovery EMAIL': 'recovery email'}
    frame = frame.rename(new_cols, axis='columns')

    # Filter out NON-"@converse.edu" addresses in "email address"
    frame.loc[:, 'email address'].replace(to_replace=r"$(?<!@converse.edu)", value=None, inplace=True, regex=True)
    # Filter out "@converse.edu" addresses in "school email"
    frame.loc[:, 'school email'].replace(to_replace=r"@converse.edu$", value=None, inplace=True, regex=True)

    # If there are recovery email addresses, use them to fill in blanks in the school emails
    if 'recovery email' in frame.columns.to_list():
        frame.loc[:, 'school email'].fillna(frame.loc[:, 'recovery email'], inplace=True)

    # Ensure names are capitalized
    for field in ['First name', 'Last name']:
        frame.loc[:, field] = frame.loc[:, field].map(cap_first, na_action='ignore')

    if Path(csvfile).is_file():
        # Read that, and update frame with the data
        frame2 = pd.read_csv(csvfile,dtype="str")
        frame2 = frame2.set_index('Jenzabar ID')
        frame2 = frame2.rename(new_cols, axis='columns')
        frame = frame.combine_first(frame2)

    # print('Returning:\n', frame, '\n')

    return frame

def send_emails(smtp: SMTP, course_dict: dict[str, str | pd.DataFrame]) -> None:
    """Takes an SMTP connection SMTP and a dict COURSE_DICT and uses them to
        send welcome emails to the students listed in COURSE_DICT['students']
        who have not already received emails.  COURSE_DICT['students'] is
        updated with the information that the new emails have been sent."""
    coursenum = course_dict['display_num'] + '.' + course_dict['section']
    lastname = course_dict['instructor'].at['lastname']                                     # type: ignore
    students = course_dict['students'].copy()                                               # type: ignore
    # print('\n', students, '\n')                                                                   # type: ignore
    # Filter out the folks who have already been sent mail
    assert 'sent' in students.columns.to_list(), 'No SENT column'                           # type: ignore
    print('Filtering students already sent email')
    sentvec = students.loc[:, "sent"].isna()
    # print(sentvec, '\n')
    students = students.loc[sentvec]                                                        # type: ignore
    print('Filtering students with no remaining emails')
    students = students.dropna(how='all', subset=['email address', 'school email'])
    # print(students, '\n')

    if not students.empty:
        msg = EmailMessage()
        msg['Subject'] = coursenum
        msg['From'] = from_addr
        msg['To'] = students.loc[:, ['email address']].dropna().to_dict(orient='list')['email address']  # type: ignore
        msg['CC'] = course_dict['instructor'].at['email']                                       # type: ignore
        msg['BCC'] = students.loc[:, ['school email']].dropna().to_dict(orient='list')['school email']   # type: ignore

        content = f"""\
Welcome!  You are registered for {coursenum}, {course_dict['coursename']}.  Your professor will be Dr. {course_dict['instructor'].at['firstname']} {lastname} ({msg['CC']}).

Your course is being taught through Converse's Canvas.  The simplest way to get in to Converse's Canvas is to go to https://converse.instructure.com and log in there with your Converse email and password.

If you've taken a course with Converse before, your Converse email address and password will normally be unchanged from what they were. If you have a new Converse account, you should have gotten your Converse email address and password in a separate email from Campus Technology.  If, for any reason, you don't have your address and password, you can get your Converse email and password by contacting Campus Technology at 864-596-9457 during their business hours (M-Th 8-5, F 8-1) or at helpdesk@converse.edu.

Once logged in, you should be taken to your Canvas dashboard.  On that dashboard, you should see a tile with the name of your course.  Click that tile to be taken to the course.
"""
        if date.fromisoformat(cast(str, course_dict['start_date'])) > date.today():
            content += """
If you don't see the tile before the course starts, that is not (yet) a problem.  If you're still not seeing your course by a day after it's supposed to start, please contact your professor.  If that doesn't help, contact me (Dr. Peter Brown, peter.brown@converse.edu).
"""
        content += f"""
Please remember that you can always email me (peter.brown@converse.edu) for Canvas questions. Dr. {lastname} is a better source for all other questions.

Peace,
—Peter Brown

Peter H. Brown, Ph.D.
Asst. Professor of Computer Science
Director of Distance Education
Converse University
"""
    
        msg.set_content(content)
        if True:   # Set True to echo the message
            print()
            print('To:', msg['To'])
            print('CC:', msg['CC'])
            print('BCC:', msg['BCC'])
            print('Subject:', msg['Subject'])
            print(msg.get_content())
            print()
        # smtp.send_message(msg)
        # 2025-05-26T15:50:57
        students['sent'] = datetime.now().isoformat(timespec='minutes')
        course_dict['students'].fillna(students, inplace=True) 
        print('cd[students]:\n', course_dict['students'], '\n')

   
def write_students_csv(course_dict: dict[str, str | pd.DataFrame]) -> None:
    """Takes a dict COURSE_DICT and writes COURSE_DICT['students'] to the
        appropriate CSV file."""
    course_dict['students'].to_csv(course_dict['students_csv'], quoting=csv.QUOTE_STRINGS) # type: ignore
    print('Wrote file', course_dict['students_csv'])

def move_to_processed(course_dict: dict[str, str | pd.DataFrame]) -> None:
    """Takes a Path INFILE and moves it to processed_folder."""
    # Exempt CSC 000 from copying
    if cast(str, course_dict['display_num']) != 'CSC 000':
        infile = Path(cast(str, course_dict['infile']))
        outfile = processed_folder.joinpath(infile.name)                    
        infile.replace(outfile)
        print('Moved', course_dict['infile'], 'to', outfile)
    
def main(args: list[str]) -> int:
    verify_constants()
    input_files = list_input_files()
    i_df = make_frame(instructors_file, 'filename')
    cnm_df = make_frame(coursenames_file, 'course')

    with SMTP('smtp.gmail.com', 587) as smtp:
        # smtp.starttls()
        # smtp.login(from_addr, "app password goes here")


        for f in input_files:
            print('Processing', f)
            if f.stem in files_to_skip:
                continue
            course_dict: dict[str, str | pd.DataFrame] = init_course_dict(f, i_df, cnm_df)
            course_dict['students'] = read_input(f, cast(str, course_dict['students_csv']))
            send_emails(smtp, course_dict)
            write_students_csv(course_dict)
            move_to_processed(course_dict)
    return 0

if __name__ == '__main__':
    import sys
    sys.exit(main(sys.argv))

# TODO:
# x1. Filter column headings
# x2. Remove non-Converse email addresses from To:
# x3. Remove Converse email addresses from BCC:
# x4. Capitalize names
# x5. Correctly handle the two EDU courses
# 6. Actually send mail?
#    - Sending with SMTP evidently requires me to create an app password.  Not sure that's possible on a Converse account.
#      - Wait--yes, it is.  Have to see whether it works, of course...
#      - Fails with """smtplib.SMTPAuthenticationError: (535, b'5.7.8 Username and Password not accepted. For more information, go to\n5.7.8  https://support.google.com/mail/?p=BadCredentials 00721157ae682-70ca8508772sm41659537b3.82 - gsmtp')"""
#        Hmm--maybe sending programmatically from a Converse account isn't so easy after all.  Talk to CT about this.
#    - Sending without SMTP means using the Gmail API from Google Cloud, but my Converse account has Google Cloud turned off.
# 7. Break dependency between processed folder and code folder
# 8. Clean repo, move code to Dropbox