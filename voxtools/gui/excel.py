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

import sys, os.path

from PyQt5.QtWidgets import (QApplication,
                            QComboBox,
                            QHBoxLayout,
                            QLabel,
                            QTextEdit,
                            QListWidget, 
                            QMessageBox,
                            QVBoxLayout,
                            QWidget,
                            )
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtGui import QIcon

# Import voxtools excel
from voxtools import excel
#DEBUG = True
DEBUG = False

# Get the gui directory
gui_dir = os.path.dirname(__file__)
gui_logo = os.path.join(gui_dir, 'icons', 'gui_logo.png').replace('\\', '/')
gui_logo_small = os.path.join(gui_dir, 'icons', 'gui_logo_small.png').replace('\\', '/')
WindowIcon = os.path.join(gui_dir, 'icons', 'WindowIcon_winICO.ico').replace('\\', '/')

# Define the texts for the combo
combo_00 = \
r"""Prettify Tabulation report

After a tabulation job, copy the output
from the browser into an Excel document.
Just store the whole output in 1 sheet.

Save, and then drag the Excel file here.

Example:
P:\Voxmeter_python_tools\voxtools\voxtools\test_suite\shared_data\wb05.xlsx
"""

combo_01 = \
r"""Single label text classification

Example:
P:\Voxmeter_python_tools\voxtools\voxtools\test_suite\shared_data\kodning01.xlsx

Make a similar file, and then drag the Excel file here.
"""

combo_02 = \
r"""Multi label text classification

Example:
P:\Voxmeter_python_tools\voxtools\voxtools\test_suite\shared_data\multikodning01.xlsx

Make a similar file, and then drag the Excel file here.
"""

combo_texts = [combo_00, combo_01, combo_02]

class MainTableWidget(QWidget):

    def __init__(self, parent=None):
        super(MainTableWidget,self).__init__(parent)

        # Title
        self.setWindowTitle('Make Excel report')
        # Icon for taskbar
        self.setWindowIcon(QIcon(WindowIcon))

        # Size and position
        w = 350; h=800
        self.resize(w, h)
        self.move(0, 0)

        # Make overall layout
        widgetLayout = QVBoxLayout(self)

        # Create layout and widgets for methods
        widgetLayout_methods = QHBoxLayout()
        # Make label and combo widget
        methods_lbl = QLabel('Chose method:')
        methods = ['Prettify Tabulation report',
                        'Single label text classification',
                        'Multi label text classification']
        # Create and fill the combo box to choose the method
        self.method_combo = QComboBox()
        self.method_combo.addItems(methods)
        # Add connect
        self.method_index = 0
        self.method_combo.currentIndexChanged.connect(self.method_change)

        # Add widgets to methods
        widgetLayout_methods.addWidget(methods_lbl)
        widgetLayout_methods.addWidget(self.method_combo)
        #widgetLayout_methods.addStretch()
        # Add to vertical
        widgetLayout.addLayout(widgetLayout_methods)

        # Create layout and widgets for info
        widgetLayout_info_lbl = QHBoxLayout()
        # The label for the info widget
        info_lbl = QLabel('Info:')
        # Add widgets to info
        widgetLayout_info_lbl.addWidget(info_lbl)
        # Add to vertical
        widgetLayout.addLayout(widgetLayout_info_lbl)
        
        # Create layout and widget for info
        widgetLayout_info = QVBoxLayout()
        # The textbox with info
        self.info_textbox = QTextEdit()
        self.info_textbox.setFixedSize(QSize(w-20, 150))
        # Set Read only
        self.info_textbox.setReadOnly(True)
        #self.info_textbox.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        # Set text
        self.set_info_in_textbox(combo_texts[self.method_index])
        # Add widgets to info
        widgetLayout_info.addWidget(self.info_textbox)
        #widgetLayout_info.addStretch()
        # Add to vertical
        widgetLayout.addLayout(widgetLayout_info)

        # Make drop_box
        widgetLayout_drop_box = QVBoxLayout()
        # Create drop_box
        drop_box_lbl = QLabel('Drag & Drop files to the box below:')
        self.drop_box = TestListBox()
        # Add widget
        widgetLayout_drop_box.addWidget(drop_box_lbl)
        widgetLayout_drop_box.addWidget(self.drop_box)
        # Add to vertical
        widgetLayout.addLayout(widgetLayout_drop_box)
        #widgetLayout_drop_box.addStretch()        
        
        # Make layout for whole gui
        self.setLayout(widgetLayout)  

    def method_change(self, i):
        # Store the index
        self.method_index = i
        # Set text
        self.set_info_in_textbox(combo_texts[self.method_index])

        # Set method to drop_box
        self.drop_box.method_index = i
        
        #print("Items in the list are :")
        #for j in range(self.method_combo.count()):
        #    print(j, self.method_combo.itemText(j))
        #print("Current index '", i, "' selection changed '", self.method_combo.currentText(), "'")

    def set_info_in_textbox(self, text):
        self.info_textbox.setText(text)


class TestListBox(QListWidget):
    def __init__(self, parent=None):
        super(TestListBox, self).__init__(parent)
        self.setAcceptDrops(True)

        # Set background
        self.setStyleSheet("""QListWidget {
                    background-color: white;
                    background-image: url(%s);
                    background-position: center;
                    background-repeat: no-repeat;
                    }"""%gui_logo_small)

        # The initial method index
        self.method_index = 0

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

            print(self.method_index)

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
            self.ok_dialog(timeout=3, text="All files has been converted!")

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
    #ex = TestListBox()
    ex = MainTableWidget()
    ex.show()
    app.exec_()
