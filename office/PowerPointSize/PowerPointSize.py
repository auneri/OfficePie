#!/usr/bin/env python
"""Report PowerPoint slide sizes.

For help in extending this template, see https://msdn.microsoft.com/en-us/VBA/VBA-PowerPoint
"""

from __future__ import absolute_import, division, print_function

import argparse
import inspect
import os
import sys
import tempfile

from qtpy import QtCore, QtGui, QtWidgets

sys.path.insert(0, os.path.abspath(os.path.join(inspect.getfile(inspect.currentframe()), '..', '..', '..')))
from office import PowerPoint  # noqa: E402, I100, I202


class Window(QtWidgets.QWidget):

    def __init__(self, *args, **kwargs):
        super(Window, self).__init__(*args, **kwargs)

        input_select = QtWidgets.QPushButton('...')
        self.input_path = QtWidgets.QLabel('Input Presentation.pptx')
        self.progress = QtWidgets.QProgressBar()
        self.table = QtWidgets.QTableWidget()
        size = QtWidgets.QPushButton('Size / Slide')

        self.progress.setTextVisible(False)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['Size in MB', '% Size'])
        self.table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        layout = QtWidgets.QGridLayout()
        layout.addWidget(input_select, 0, 0)
        layout.addWidget(self.input_path, 0, 1)
        layout.addWidget(self.progress, 2, 0, 1, 2)
        layout.addWidget(size, 3, 0, 1, 2)
        layout.addWidget(self.table, 4, 0, 1, 2)
        layout.setColumnStretch(1, 1)
        layout.setRowStretch(4, 1)
        self.setLayout(layout)

        input_select.clicked.connect(self.on_input_select)
        size.clicked.connect(self.on_size)

        self.setAcceptDrops(True)
        self.setAutoFillBackground(True)
        self.setWindowTitle('PowerPoint Size')
        self.show()

    def dragEnterEvent(self, event):  # noqa: N802
        if event.mimeData().urls() and event.mimeData().urls()[0].toLocalFile().endswith('.pptx'):
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
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select input presentation', self.input_path.text(), 'PowerPoint Presentations (*.pptx)')
        if path:
            self.input_path.setText(os.path.abspath(path))

    def on_size(self):
        p = PowerPoint(self.input_path.text())
        self.progress.setMaximum(p.doc.Slides.Count)
        self.table.setRowCount(p.doc.Slides.Count)
        f = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
        f.close()
        sizes = []
        for i in range(p.doc.Slides.Count):
            self.progress.setValue(i)
            QtCore.QCoreApplication.processEvents(QtCore.QEventLoop.AllEvents, 100)
            p.export(f.name, i + 1)
            sizes.append(os.path.getsize(f.name) / 1e6)
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem('{:.2f}'.format(sizes[i])))
        os.remove(f.name)
        self.progress.setValue(0)
        for i, size in enumerate(sizes):
            self.table.setItem(i, 1, QtWidgets.QTableWidgetItem('{:.0f}'.format(100 * sizes[i] / sum(sizes))))
        p.close(alert=False)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description='Reports PowerPoint slide sizes')
        parser.add_argument('input', help='Input presentation')
        args = parser.parse_args()

        p = PowerPoint(args.input)
        f = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
        f.close()
        for i in range(1, p.doc.Slides.Count + 1):
            p.export(f.name, i)
            print('{:>3}/{}: {:.1f} MB'.format(i, p.doc.Slides.Count, os.path.getsize(f.name) / 1e6))
        os.remove(f.name)
        del p
    else:
        app = QtWidgets.QApplication(sys.argv)
        window = Window()
        sys.exit(app.exec_())
