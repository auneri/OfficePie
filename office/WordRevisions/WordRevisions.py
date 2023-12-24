#!/usr/bin/env python
"""Convert Word tracked changes to formatted text.

For help in extending this template, see https://msdn.microsoft.com/en-us/VBA/VBA-Word
"""

import argparse
import os
import sys

import office
from PyQt5 import QtCore, QtGui, QtWidgets


class Window(QtWidgets.QWidget):

    def __init__(self, *args, **kwargs):
        super(Window, self).__init__(*args, **kwargs)

        input_select = QtWidgets.QPushButton('...')
        output_select = QtWidgets.QPushButton('...')
        self.input_path = QtWidgets.QLabel('Input Document.docx')
        self.output_path = QtWidgets.QLabel(os.path.abspath(os.path.expanduser('~/Desktop/Output Document.docx')))
        self.strike_deletions = QtWidgets.QCheckBox('Strike Deletions')
        self.progress = QtWidgets.QProgressBar()
        mark = QtWidgets.QPushButton('Mark')

        self.progress.setTextVisible(False)

        layout = QtWidgets.QGridLayout()
        layout.addWidget(input_select, 0, 0)
        layout.addWidget(self.input_path, 0, 1)
        layout.addWidget(output_select, 1, 0)
        layout.addWidget(self.output_path, 1, 1)
        layout.addWidget(self.strike_deletions, 2, 0, 1, 2)
        layout.addWidget(self.progress, 3, 0, 1, 2)
        layout.addWidget(mark, 4, 0, 1, 2)
        layout.setColumnStretch(1, 1)
        self.setLayout(layout)

        input_select.clicked.connect(self.on_input_select)
        output_select.clicked.connect(self.on_output_select)
        mark.clicked.connect(self.on_mark)

        self.setAcceptDrops(True)
        self.setAutoFillBackground(True)
        self.setWindowTitle('Mark Revisions')
        self.show()

    def dragEnterEvent(self, event):  # noqa: N802
        if event.mimeData().urls() and event.mimeData().urls()[0].toLocalFile().endswith('.docx'):
            self.setBackgroundRole(QtGui.QPalette.Highlight)
            event.accept()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):  # noqa: N802
        self.setBackgroundRole(QtGui.QPalette.Window)
        event.accept()

    def dropEvent(self, event):  # noqa: N802
        self.setBackgroundRole(QtGui.QPalette.Window)
        path = event.mimeData().urls()[0].toLocalFile()
        self.input_path.setText(os.path.abspath(path))
        event.accept()

    def on_input_select(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select input document', self.input_path.text(), 'Word Documents (*.docx)')
        if path:
            self.input_path.setText(os.path.abspath(path))

    def on_output_select(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Select output document', self.output_path.text(), 'Word Documents (*.docx)')
        if path:
            self.output_path.setText(os.path.abspath(path))

    def on_mark(self):
        doc = office.Word(self.input_path.text())
        self.progress.setMaximum(doc.doc.Revisions.Count)
        for n in doc.mark_revisions(strike_deletions=self.strike_deletions.isChecked()):
            self.progress.setValue(n)
            QtCore.QCoreApplication.processEvents(QtCore.QEventLoop.AllEvents, 100)
        doc.doc.SaveAs(self.output_path.text())
        doc.close(alert=False)
        self.progress.setValue(0)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description='Converts Word tracked changes to formatted text')
        parser.add_argument('input', help='Input document')
        parser.add_argument('output', help='Output document')
        parser.add_argument('--strike-deletions', action='store_true', help='Strike deletions instead of removing them')
        args = parser.parse_args()

        doc = office.Word(args.input)
        N = doc.doc.Revisions.Count
        for n in doc.mark_revisions(args.strike_deletions):
            sys.stdout.write(f'\rMarking... {100 * n / N:.0f}%')
            sys.stdout.flush()
        doc.doc.SaveAs(args.output)
        del doc
    else:
        app = QtWidgets.QApplication(sys.argv)
        window = Window()
        sys.exit(app.exec_())
