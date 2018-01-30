################################################################################
# Copyright (C) Voxmeter A/S - All Rights Reserved
#
# Voxmeter A/S
# Borgergade 6, 4.
# 1300 Copenhagen K
# Denmark
#
# Written by Troels Schwarz-Linnet <tsl@voxmeter.dk>, 2018
# 
# Unauthorized copying of this file, via any medium is strictly prohibited.
#
# Any use of this code is strictly unauthorized without the written consent
# by Voxmeter A/S. This code is proprietary of Voxmeter A/S.
# 
################################################################################

import datetime, os, os.path, tempfile, unittest

# Import voxtools sklearn_multilabel
from voxtools import sklearn_multilabel


class Test_multikodning01(unittest.TestCase):
    def setUp(self):
        # Get the time
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")
        # Create file obj
        file_obj = tempfile.NamedTemporaryFile(delete=True)
        self.excel_dst = file_obj.name + ".xlsx"
        # Get the folder with shared_data
        cur_dir = os.path.dirname(__file__)
        shared_data_dir = os.path.join(cur_dir, 'shared_data')
        # Get the filepath
        self.excel_src = os.path.join(shared_data_dir, 'multikodning01.xlsx')

        # Instantiate the Excel class
        self.tcl = sklearn_multilabel.text(excel_src=self.excel_src, excel_dst=self.excel_dst)
        # Run it
        self.tcl.run_all()

    def test_wb(self):
        # Assert the sheetnames
        self.assertEqual(self.tcl.wb.sheetnames, ['classification', 'target_categories'])

    def tearDown(self):
        # Delete temporary file
        print("Deleting : %s"%self.excel_dst)
        os.remove(self.excel_dst)


if __name__ == '__main__':
    unittest.main()