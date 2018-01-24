########################################################################################################
# From: http://mgraessle.net/2013/07/07/pyqt-and-drag-and-drop-functionality/
#
# 2018-01-24 By Troels Schwarz-Linnet

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