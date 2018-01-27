# From: http://mgraessle.net/2013/07/07/pyqt-and-drag-and-drop-functionality/

import sys, time, os
 
from PyQt5.QtWidgets import (QListWidget, QApplication)
from PyQt5.QtCore import Qt
 
class TestListBox(QListWidget):
    def __init__(self, parent=None):
        super(TestListBox, self).__init__(parent)
        self.setAcceptDrops(True)

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

            links = []
            for url in event.mimeData().urls():
                links.append(str(url.toLocalFile()))
            print(links)
            # self.emit(ui.QtCore.SIGNAL("dropped"), links)

        else:
            event.ignore()

if __name__ == '__main__':
 
    app = QApplication(sys.argv)
    ex = TestListBox()
    ex.show()
    app.exec_()