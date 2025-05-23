import os
import unittest
from welcome_students import *
# import the code you want to test here

class TestNothing(unittest.TestCase):

    # Every method that starts with the string "test"
    # will be executed as a unit test
    def testDropboxDir(self) -> None:
        self.assertEqual(str(dropbox_dir), os.path.expanduser('~/Dropbox'))

    def testInboxDir(self) -> None:
        self.assertEqual(str(dropbox_dir), os.path.expanduser('~/Dropbox'))



if __name__ == '__main__':
    unittest.main()

