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

from textblob.classifiers import NaiveBayesClassifier


class text:
    def __init__(self, excel_src=None, excel_dst=None):
        # Store 
        self.excel_src = excel_src
        self.excel_dst = excel_dst

        # Make current time
        #self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W-H%H-M%M-S%S")
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")

    def run_all(self):
        # Copy Excel workbook
        self.copy_excel()
        # Print all objects of workbook
        #[print(i) for i in dir(self.wb)]
        #print(self.wb.active, self.wb.properties)

        # Open workbook
        self.load_wb()

        # Delete empty sheets
        self.delete_empty_sheets()

        # Read data
        self.read_data()

        # Now train
        self.do_train()

        # Now classify all sentences
        self.do_classify_sentences()


    def do_classify_sentences(self):
        sentences = self.sent_dat[1]
        
        # Loop over all sentences
        self.sent_classifications = []
        for sent in sentences:
            # Do classification
            classification = self.cl.classify(sent)
            # Store
            self.sent_classifications.append((sent, classification))

    def do_train(self):
        # Get the data
        train = self.train_dat[1]

        # http://textblob.readthedocs.io/en/dev/classifiers.html#classifying-text
        # Now we’ll create a Naive Bayes classifier, passing the training data into the constructor.
        self.cl = NaiveBayesClassifier(train)

        # If we have test data, we can score the classifier
        if self.has_test:
            test = self.test_dat[1]
            self.accuracy = self.cl.accuracy(test)
            print("After training %i sentences, the accuracy is %0.2f"% (len(train), self.accuracy) )

    def read_data(self):
        # Use the first sheet
        ws_name = self.wb.sheetnames[0]
        ws = self.wb[ws_name]
        
        # Get row 1
        self.header_vals = []
        for iCol in range(1, ws.max_column):
            iCol_L = get_column_letter(iCol)
            cCell = ws['%s1'%(iCol_L)]
            self.header_vals.append(cCell.value)
        # Check for header
        if 'train' in self.header_vals:
            self.has_header = True
            self.iRow_skip = 1
        else:
            self.has_header = False
            #self.header_vals = ['sentences', 'train', 'classify', 'test', 'correct_train', 'correct_test']
            self.header_vals = ['sentences', 'train', 'classify']
            self.iRow_skip = 0

        # Get the index of columns
        index_sent = self.header_vals.index('sentences')
        index_train = self.header_vals.index('train')
        index_class = self.header_vals.index('classify')

        # Collect all sentences
        self.sent_dat = self.collect_column_data(ws, index_sent)

        # Collect paired values of sentences and train
        self.train_dat = self.collect_column_data(ws, index_sent, index_train)

        # If test is available
        if 'test' in self.header_vals:
            self.has_test = True
            index_test = self.header_vals.index('test')
            self.test_dat = self.collect_column_data(ws, index_sent, index_test)
        else:
            self.has_test = False


    def collect_column_data(self, ws, col_i_1=1, col_i_2=None):
        # Loop over rows
        iRow_nrs = []
        iRow_pair_vals = []
        
        for iRow, cCol in enumerate(ws.rows):
            # Possible skip
            if iRow < self.iRow_skip:
                continue
            if col_i_2 != None:
                # Get the cells
                cCell_1 = cCol[col_i_1]
                cCell_2 = cCol[col_i_2]
                # Get the value
                val_1 = cCell_1.value
                val_2 = cCell_2.value
                if val_1 != None and val_2 != None:
                    iRow_nrs.append((iRow,col_i_1,col_i_2))
                    iRow_pair_vals.append((val_1, val_2))

            elif col_i_2 == None:
                # Get the cells
                cCell_1 = cCol[col_i_1]
                # Get the value
                val_1 = cCell_1.value
                if val_1 != None:
                    iRow_nrs.append((iRow,col_i_1))
                    iRow_pair_vals.append(val_1)

        if len(iRow_nrs) > 0:
            return [iRow_nrs, iRow_pair_vals]
        else:
            return None

    def delete_empty_sheets(self):
        # Loop through worksheets
        for ws in self.wb:
            #print(ws.title, ws.max_row, ws.max_column)
            # Delete sheet if empty
            if ws.max_row == 1 and ws.max_column == 1:
                #print(ws.title)
                std=self.wb[ws.title]
                self.wb.remove(std)

    def load_wb(self):
        # Open workbook
        self.wb = load_workbook(self.excel_dst)

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
