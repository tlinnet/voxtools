################################################################################
# Copyright (C) Troels Schwarz-Linnet - All Rights Reserved
# Written by Troels Schwarz-Linnet <tlinnet@gmail.com>, January 2018
# 
# Unauthorized copying of this file, via any medium is strictly prohibited.
#
# Any use of this code is strictly unauthorized without the written consent
# by the the author. This code is proprietary of the author.
# 
################################################################################

import datetime, os.path, tempfile, unittest

# Import voxtools excel
from voxtools import excel

class Test_excel(unittest.TestCase):
    def setUp(self):
        # Get the time
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")

    def test_copy_excel(self):
        # Create file obj
        file_obj = tempfile.NamedTemporaryFile(delete=True)
        file_name = file_obj.name

        # Instantiate the Excel class
        exl = excel.excel(file_name)
        # Run the class function
        exl.copy_excel()

        # Check that the new file has been created
        new_file_name = file_name + "_" +  self.cur_time
        file_exists = os.path.isfile(new_file_name)
        # Assert this is true
        self.assertTrue(file_exists)

        # Instantiate a new Excel class
        file_name_new = file_name + "_test"
        exl = excel.excel(file_name, file_name_new)
        # Run the class function
        exl.copy_excel()
        file_exists = os.path.isfile(file_name_new)
        #print(file_name_new)
        self.assertTrue(file_exists)



if __name__ == '__main__':
    unittest.main()