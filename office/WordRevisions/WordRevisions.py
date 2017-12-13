#!/usr/bin/env python

# TODO Allow marking individual authors.
# TODO Customize color.
# TODO Allow use of theme colors.

"""
To create a portable application, run:
    pyinstaller MarkRevisions.spec
"""

from __future__ import absolute_import, division, print_function

import argparse
import inspect
import os
import sys

import qtpy
from qtpy import QtCore, QtGui, QtWidgets

from office import Word  # noqa: E402, I100, I202

__author__ = 'Ali Uneri'

os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.abspath(os.path.join(inspect.getfile(qtpy), '..', '..', 'PyQt5', 'Qt', 'plugins'))


class Window(QtWidgets.QWidget):

    def __init__(self, *args, **kwargs):
        super(Window, self).__init__(*args, **kwargs)

        input_select = QtWidgets.QPushButton('...')
        output_select = QtWidgets.QPushButton('...')
        self.input_path = QtWidgets.QLabel('input.docx')
        self.output_path = QtWidgets.QLabel(os.path.abspath(os.path.expanduser('~/Desktop/output.docx')))
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

        input_select.clicked.connect(self.on_input_select)
        output_select.clicked.connect(self.on_output_select)
        mark.clicked.connect(self.on_mark)

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
        w = Word(self.input_path.text())
        self.progress.setMaximum(w.doc.Revisions.Count)
        for n in w.mark_revisions(strike_deletions=self.strike_deletions.isChecked()):
            self.progress.setValue(n)
            QtCore.QCoreApplication.processEvents(QtCore.QEventLoop.AllEvents, 100)
        w.doc.SaveAs(self.output_path.text())
        w.close(alert=False)
        self.progress.setValue(0)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description='Convert tracked changes to marked revisions')
        parser.add_argument('input', help='Input document')
        parser.add_argument('output', help='Output document')
        parser.add_argument('--strike-deletions', action='store_true', help='Strike deletions instead of removing them')
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
