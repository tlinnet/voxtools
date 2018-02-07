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

# Import voxtools excel
from voxtools import ascii_mod

class Test_excel(unittest.TestCase):
    def setUp(self):
        # Get the time
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")

        # Get the folder with shared_data
        cur_dir = os.path.dirname(__file__)
        shared_data_dir = os.path.join(cur_dir, 'shared_data')
        # Get the filepath
        self.ascii_f = os.path.join(shared_data_dir, 'ascii_def.txt')

    def test_class(self):
        # Instantiate the class
        self.asc =  ascii_mod.create_ascii_input(ascii_f=self.ascii_f)

        
        


if __name__ == '__main__':
    unittest.main()