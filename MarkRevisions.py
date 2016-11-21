#!/usr/bin/env python

"""
To create a portable application, run
pyinstaller --clean --name=MarkRevisions --onefile --windowed --icon=MarkRevisions.ico MarkRevisions.py
"""

from __future__ import absolute_import, division, print_function

import argparse
import os
import sys

from qtpy import QtGui, QtWidgets

from PythonTools.helpers.Office import Word

__author__ = 'Ali Uneri'


class Window(QtWidgets.QWidget):

    def __init__(self, parent=None):
        super(Window, self).__init__(parent)

        input_select = QtWidgets.QPushButton('...')
        output_select = QtWidgets.QPushButton('...')
        self.input_path = QtWidgets.QLabel('input.docx')
        self.output_path = QtWidgets.QLabel(os.path.normpath(os.path.expanduser('~/Desktop/output.docx')))
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
        layout.setColumnStretch(1,1)
        self.setLayout(layout)

        input_select.clicked.connect(self.set_input)
        output_select.clicked.connect(self.set_output)
        mark.clicked.connect(self.mark)

        self.setAcceptDrops(True)
        self.setAutoFillBackground(True)
        self.setWindowTitle('Mark Revisions')
        self.show()

    def dragEnterEvent(self, event):
        if event.mimeData().urls() and event.mimeData().urls()[0].toLocalFile().endswith('.docx'):
            self.setBackgroundRole(QtGui.QPalette.Highlight)
            event.accept()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setBackgroundRole(QtGui.QPalette.Window)
        event.accept()

    def dropEvent(self, event):
        self.setBackgroundRole(QtGui.QPalette.Window)
        path = event.mimeData().urls()[0].toLocalFile()
        self.input_path.setText(os.path.normpath(path))
        event.accept()

    def set_input(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select input document', self.input_path.text(), 'Word Documents (*.docx)')
        if path:
            self.input_path.setText(os.path.normpath(path))

    def set_output(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Select output document', self.output_path.text(), 'Word Documents (*.docx)')
        if path:
            self.output_path.setText(os.path.normpath(path))

    def mark(self):
        w = Word(self.input_path.text())
        self.progress.setMaximum(w.doc.Revisions.Count)
        for n in w.mark_revisions(self.strike_deletions.isChecked()):
            self.progress.setValue(n)
        w.doc.SaveAs(self.output_path.text())
        self.progress.setValue(0)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description='Convert tracked changes to marked revisions')
        parser.add_argument('input', help='Input document')
        parser.add_argument('output', help='Output document')
        parser.add_argument('-sd', '--strike-deletions', nargs='?', const=True, default=False, type=int, help='Strike deletions instead of removing them')
        args = parser.parse_args()

        w = Word(args.input)
        N = w.doc.Revisions.Count
        for n in w.mark_revisions(args.strike_deletions):
            sys.stdout.write('\rMarking... {:.0f}%'.format(100 * n / N))
            sys.stdout.flush()
        w.doc.SaveAs(args.output)
    else:
        app = QtWidgets.QApplication(sys.argv)
        window = Window()
        sys.exit(app.exec_())
