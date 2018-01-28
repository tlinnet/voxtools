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

# Import voxtools textblob_classifying
from voxtools import textblob_classifying

class Test_excel(unittest.TestCase):
    def setUp(self):
        # Get the time
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")

    def test_copy_excel(self):
        # Create file obj
        file_obj = tempfile.NamedTemporaryFile(delete=True)
        excel_src = file_obj.name
        file_dir = os.path.dirname(excel_src)

        # Instantiate the Text Class
        tcl = textblob_classifying.text(excel_src)
        # Run the class function
        tcl.copy_excel()

        # Check that the new file has been created
        excel_src_new = excel_src + "_" +  self.cur_time
        file_exists = os.path.isfile(excel_src_new)
        # Assert this is true
        self.assertTrue(file_exists)

        # Instantiate a new Excel class, and reuse name
        excel_src_new_update_1 = excel_src_new + "_002"
        tcl = textblob_classifying.text(excel_src_new)
        # Run the class function
        tcl.copy_excel()
        file_exists = os.path.isfile(excel_src_new_update_1)
        #print("%s\n%s"%(excel_src_new, excel_src_new_update_1))
        self.assertTrue(file_exists)
        # Delete previous
        os.remove(excel_src_new)

        # Instantiate a new Excel class, and reuse name
        excel_src_new_update_2 = excel_src_new + "_003"
        tcl = textblob_classifying.text(excel_src_new_update_1)
        # Run the class function
        tcl.copy_excel()
        file_exists = os.path.isfile(excel_src_new_update_2)
        #print("%s\n%s"%(excel_src_new_update_1, excel_src_new_update_2))
        self.assertTrue(file_exists)
        # Delete previous
        os.remove(excel_src_new_update_1)
        os.remove(excel_src_new_update_2)


if __name__ == '__main__':
    unittest.main()