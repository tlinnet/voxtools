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

import sys, os.path

from PyQt5.QtWidgets import (QApplication, QListWidget, QMessageBox)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QIcon

# Import voxtools excel
from voxtools import excel
#DEBUG = True
DEBUG = False

# Get the gui directory
gui_dir = os.path.dirname(__file__)
gui_logo = os.path.join(gui_dir, 'icons', 'gui_logo.png').replace('\\', '/')
WindowIcon = os.path.join(gui_dir, 'icons', 'WindowIcon_winICO.ico').replace('\\', '/')

class TestListBox(QListWidget):
    def __init__(self, parent=None):
        super(TestListBox, self).__init__(parent)
        self.setAcceptDrops(True)

        # Title
        self.setWindowTitle('Make Excel report')
        # Icon for taskbar
        self.setWindowIcon(QIcon(WindowIcon))
        # Size and position
        w = 300; h=600
        self.resize(w, h)
        self.move(0, 0)
        # Set background
        self.setStyleSheet("""QListWidget {
                    background-color: white;
                    background-image: url(%s);
                    background-position: center;
                    background-repeat: no-repeat;
                    }"""%gui_logo)


    # http://pyqt.sourceforge.net/Docs/PyQt5/api/QtGui/qdragenterevent.html
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
              event.accept()
        else:
              event.ignore()

    # http://pyqt.sourceforge.net/Docs/PyQt5/api/QtGui/qdragmoveevent.html
    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
              event.setDropAction(Qt.CopyAction)
              event.accept()
        else:
              event.ignore()

    # http://pyqt.sourceforge.net/Docs/PyQt5/api/QtGui/qdropevent.html
    def dropEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()

            # Loop over urls passed
            for url in event.mimeData().urls():
                link_str = str(url.toLocalFile())
                # Get filename_src
                filename_src, fileext = os.path.splitext(link_str)
                # If not Excel
                if fileext != '.xlsx':
                    print("Not Excel file! Skipping file: %s" % link_str)
                    continue
                # If Excel
                # Instantiate the Excel class
                if not DEBUG:
                    if "kodning" in link_str.lower():
                        from voxtools import textblob_classifying
                        tcl = textblob_classifying.text(excel_src=link_str)
                        # Run it
                        tcl.run_all()

                    else:
                        exl = excel.excel(excel_src=link_str)
                        # Run it
                        exl.run_all()

            # Show dialog
            self.ok_dialog(timeout=3, text="All Excel files has been converted!")

        else:
            event.ignore()

    def ok_dialog(self, timeout=3, text=None):
        messagebox = TimerMessageBox(timeout=timeout, text=text, parent=self)
        messagebox.exec_()

    def showdialog(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)

        msg.setText("This is a message box")
        msg.setInformativeText("This is additional information")
        msg.setWindowTitle("MessageBox demo")
        msg.setDetailedText("The details are as follows:")
        msg.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg.buttonClicked.connect(self.msgbtn)

        retval = msg.exec_()
        print("value of pressed message box button:", retval)

    def msgbtn(self,i):
        print("Button pressed is:",i.text())

class TimerMessageBox(QMessageBox):
    def __init__(self, timeout=3, text=None, parent=None):
        super(TimerMessageBox, self).__init__(parent)
        self.setWindowTitle("Ok status")
        self.time_to_wait = timeout
        self.text = text
        if self.text == None:
            self.setText("wait (closing automatically in {0} secondes.)".format(timeout))
        else:
            self.setText("{0}\n\n      Window closes in {1}s".format(self.text, self.time_to_wait))
        self.setStandardButtons(QMessageBox.NoButton)
        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.changeContent)
        self.timer.start()

    def changeContent(self):
        if self.text == None:
            self.setText("wait (closing automatically in {0} secondes.)".format(self.time_to_wait))
        else:
            self.setText("{0}\n\n      Window closes in {1}s".format(self.text, self.time_to_wait))
        self.time_to_wait -= 1
        if self.time_to_wait <= 0:
            self.close()

    def closeEvent(self, event):
        self.timer.stop()
        event.accept()

if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = TestListBox()
    ex.show()
    app.exec_()
