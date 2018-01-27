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

import sys, os, datetime, shutil
 
from PyQt5.QtWidgets import (QListWidget, QApplication)
from PyQt5.QtCore import Qt
 
class TestListBox(QListWidget):
    def __init__(self, parent=None):
        super(TestListBox, self).__init__(parent)
        self.setAcceptDrops(True)

        # Make current time
        self.cur_time = datetime.datetime.now().strftime("%Y-%m-%d-Week%W")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
              event.accept()
        else:
              event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
              event.setDropAction(Qt.CopyAction)
              event.accept()
        else:
              event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()

            for url in event.mimeData().urls():
                link_str = str(url.toLocalFile())
                # Get filename_src
                filename_src, fileext = os.path.splitext(link_str)
                # If not Excel
                if fileext != '.xlsx':
                    print("Not Excel file! Skipping file: %s" % link_str)
                    continue
                # If Excel
                filename_dst = filename_src + "_" +  self.cur_time
                # New destination
                dst = filename_dst+fileext
                # Copy
                shutil.copy2(link_str, dst)

            #print(links)
            # self.emit(ui.QtCore.SIGNAL("dropped"), links)

        else:
            event.ignore()

if __name__ == '__main__':
 
    app = QApplication(sys.argv)
    ex = TestListBox()
    ex.show()
    app.exec_()