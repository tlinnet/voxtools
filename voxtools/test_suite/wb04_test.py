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
from openpyxl.utils import get_column_letter

# Import voxtools excel
from voxtools import excel


class Test_wb04(unittest.TestCase):
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
        self.excel_src = os.path.join(shared_data_dir, 'wb04.xlsx')

        # Instantiate the Excel class
        self.exl = excel.excel(excel_src=self.excel_src, excel_dst=self.excel_dst)
        # Copy Excel workbook
        self.exl.copy_excel()
        # Open workbook
        self.exl.load_wb()
        # Delete empty sheets
        self.exl.delete_empty_sheets()
        # Format number percent format per cell
        self.exl.Format_Cells_number_format()

    def test_make_uniq_key(self):
        # Test the creation of new keys
        keylist = ['a', 'b']
        key = 'c'

        # Assert normal mode
        new_key = self.exl.make_uniq_key(key, keylist)
        self.assertEqual(new_key, key)

        # Change keylist
        keylist = ['a', 'b', 'c']
        new_key = self.exl.make_uniq_key(key, keylist)
        self.assertEqual(new_key, 'c_2')

        # Change key
        keylist = ['a', 'b', 'c', 'c_2']
        new_key = self.exl.make_uniq_key(key, keylist)
        self.assertEqual(new_key, 'c_3')


    def test_Find_sheet_groups(self):
        # Find the Sheet groups
        self.exl.Find_sheet_groups()

        # Check that keys are uniq
        keys = self.exl.sheet_groups_dict['Ark1']['keys']
        len_keys = len(keys)
        len_keys_set = len(set(keys))
        self.assertEqual(len_keys, len_keys_set)

    def test_run_all(self):
        # Find the Sheet groups
        self.exl.Find_sheet_groups()
        # Copy groups to sheets
        self.exl.Copy_groups_to_sheets_as_copy()
        # Set width of cells
        self.exl.Format_Column_Width()
        # Set Page Layout per sheet
        self.exl.Page_Layout_setup()
        # Set height of cells
        self.exl.Format_Column_Height()
        # Save Workbook
        self.exl.wb.save(self.excel_dst)


    def tearDown(self):
        # Delete temporary file
        print("Deleting : %s"%self.excel_dst)
        os.remove(self.excel_dst)

if __name__ == '__main__':
    unittest.main()