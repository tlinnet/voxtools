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
import shutil, datetime, os, copy, operator, math

from openpyxl import load_workbook
from openpyxl.utils import coordinate_from_string, column_index_from_string, get_column_letter
#from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles import Font, PatternFill, Border

# Print version
#from openpyxl import __version__; print(__version__)

# Set standards

class text:
    def __init__(self, excel_src=None, excel_dst=None):
        # Store 
        self.excel_src = excel_src
        self.excel_dst = excel_dst

        # Make current time
        #self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W-H%H-M%M-S%S")
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")


    def copy_excel(self):
        filename_src, fileext = os.path.splitext(self.excel_src)
        filename_dst = filename_src + "_" +  self.cur_time
        # New destination
        if self.excel_dst == None:
            # If this is an update on file with date in it
            if self.cur_time in filename_src:
                filename_split = filename_src.split(self.cur_time)
                if filename_split[-1] == "":
                    filename_dst = filename_split[0] + self.cur_time + "_002"
                else:
                    version = int(filename_split[-1].split("_")[-1])
                    version += 1
                    filename_dst = filename_split[0] + self.cur_time + "_%03d"%(version)
            
            self.excel_dst = filename_dst+fileext

        # Copy
        shutil.copy2(self.excel_src, self.excel_dst)
