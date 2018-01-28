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

import datetime, os, os.path, tempfile, unittest

# Import voxtools excel
from voxtools import excel

class Test_excel(unittest.TestCase):
    def setUp(self):
        # Get the time
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")

    def test_copy_excel(self):
        # Create file obj
        file_obj = tempfile.NamedTemporaryFile(delete=True)
        excel_src = file_obj.name

        # Instantiate the Excel class
        exl = excel.excel(excel_src)
        # Run the class function
        exl.copy_excel()

        # Check that the new file has been created
        excel_src_new = excel_src + "_" +  self.cur_time
        file_exists = os.path.isfile(excel_src_new)
        # Assert this is true
        self.assertTrue(file_exists)
        # Delete 
        os.remove(excel_src_new)

        # Instantiate a new Excel class
        excel_src_new = excel_src + "_test"
        exl = excel.excel(excel_src, excel_src_new)
        # Run the class function
        exl.copy_excel()
        file_exists = os.path.isfile(excel_src_new)
        #print(excel_src_new)
        self.assertTrue(file_exists)
        # Delete
        os.remove(excel_src_new)


if __name__ == '__main__':
    unittest.main()