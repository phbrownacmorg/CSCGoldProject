import os
import pandas as pd
import unittest
from smtplib import SMTP
from typing import cast
from welcome_students import *
# import the code you want to test here

class TestWelcome(unittest.TestCase):

    def setUp(self) -> None:
        self._CSC000stem = 'CSC000.00_2025-06-01_2425-BS_Brown,P.'
        self._CSC000fullname = self._CSC000stem[:9] + '-' + self._CSC000stem[21:28]
        self._CSC000xlsx_path = input_folder.joinpath(self._CSC000stem + '.xlsx')
        self._PLP700stem = 'PLP700.Y1_2025-05-19_2425-BS_Gilliam,K.'
        self._PLP700fullname = self._PLP700stem[:9] + '-' + self._PLP700stem[21:28]
        self._PLP700xlsx_path = input_folder.joinpath(self._PLP700stem + '.xlsx')
        self._PLP701stem = 'PLP701.Y1_2025-05-19_2425-BS_Knipe,J.'
        self._PLP701fullname = self._PLP701stem[:9] + '-' + self._PLP701stem[21:28]
        self._PLP701xlsx_path = input_folder.joinpath(self._PLP701stem + '.xlsx')
        self._PLP721stem = 'PLP721.Y1_2025-05-27_2425-BS_Thomas,E.'

        self._inst_f = make_frame(instructors_file, 'filename')
        self._crs_f = make_frame(coursenames_file, 'course')

        self._CSC000_d = init_course_dict(self._CSC000xlsx_path, self._inst_f, self._crs_f)
        self._CSC000_d['students'] = read_input(self._CSC000xlsx_path, cast(str, self._CSC000_d['students_csv']))

    # Every method that starts with the string "test"
    # will be executed as a unit test
    def testDropboxDir(self) -> None:
        self.assertEqual(str(dropbox_dir), os.path.expanduser('~/Dropbox'))
        self.assertTrue(dropbox_dir.is_dir())

    def testInboxDir(self) -> None:
        self.assertEqual(str(input_folder), os.path.expanduser('~/Dropbox/DEd/students_in'))
        self.assertTrue(input_folder.is_dir())

    def testCodeDir(self) -> None:
        self.assertEqual(str(code_folder), os.path.dirname(__file__))
        self.assertTrue(code_folder.is_dir())

    def testProcessedDir(self) -> None:
        self.assertEqual(str(processed_folder), os.path.join(os.path.dirname(__file__), 'students_processed'))
        self.assertTrue(processed_folder.is_dir())

    def testCSVDir(self) -> None:
        self.assertEqual(str(csv_folder), os.path.join(os.path.dirname(__file__), 'students_csv'))
        self.assertTrue(csv_folder.is_dir())

    def testVerifyConstants(self) -> None:
        self.assertTrue(verify_constants())

    # Only works at home in May 2025.  Skip otherwise!
    @unittest.skip('')
    def testListFiles(self) -> None:
        self.assertEqual(list_input_files(),
                         [self._CSC000xlsx_path, input_folder.joinpath('EDU591.Y2_anonymized_2425-AS_Bradley,C..xlsx'),
                          input_folder.joinpath('EDU592.Y3_anonymized_2425-AS_Bradley,C..xlsx'),
                          self._PLP700xlsx_path, self._PLP701xlsx_path,
                          input_folder.joinpath(self._PLP721stem + '.xlsx')])
                          
    def testInitCourseDictCSC000(self) -> None:
        d = init_course_dict(self._CSC000xlsx_path, self._inst_f, self._crs_f)
        self.assertEqual(d['infile'], str(self._CSC000xlsx_path))
        self.assertEqual(d['prefix'], 'CSC')
        self.assertEqual(d['barenum'], '000')
        self.assertEqual(d['section'], '00')
        self.assertEqual(d['display_num'], 'CSC 000')
        self.assertEqual(d['coursename'], 'Bogus Course')
        self.assertEqual(d['term'], '2425-BS')
        self.assertEqual(d['full_num'], self._CSC000fullname)
        self.assertEqual(d['students_csv'], os.path.join(str(csv_folder), self._CSC000fullname + '.csv'))
        self.assertEqual(cast(str, cast (pd.DataFrame, d['instructor'])['firstname']), 'Peter')
        self.assertEqual(cast(str, cast (pd.DataFrame, d['instructor'])['lastname']), 'Brown')
        self.assertEqual(cast(str, cast (pd.DataFrame, d['instructor'])['email']), 'Peter.Brown@converse.edu')
        self.assertEqual(d['start_date'], '2025-06-01')
        
    def testReadInputCSC000_noCSV(self) -> None:
        frame = read_input(self._CSC000xlsx_path, os.path.join(csv_folder, self._CSC000fullname + '.csv'))
        print('\ntestReadInputCSC000:\n', frame, '\n')
        #print(frame.loc['1037970'])

    @unittest.skip('')
    def testWriteCSV_CSC000(self) -> None:
        write_students_csv(self._CSC000_d)

    def testPrintStudents(self) -> None:
        with SMTP('smtp.gmail.com', 587) as smtp:
            send_emails(smtp, self._CSC000_d)

if __name__ == '__main__':
    unittest.main()

# Notes, 2025-05-22 AM: I suspect this may not work for an email account that has 2FA.  Even so, having the email
# preprocessed out for cut-and-paste may have advantages; and especially being able to track who's been emailed in
# the past, and at which addresses, may be valuable.