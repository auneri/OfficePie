#!/usr/bin/env python
"""Report PowerPoint slide sizes.

For help in extending this template, see https://msdn.microsoft.com/en-us/VBA/VBA-PowerPoint
"""

import argparse
import pathlib
import sys
import tempfile

from PyQt5 import QtCore, QtGui, QtWidgets

import office


class Window(QtWidgets.QWidget):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

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
        filepath = event.mimeData().urls()[0].toLocalFile()
        self.input_path.setText(str(pathlib.Path(filepath).resolve()))
        event.accept()

    def on_input_select(self):
        filepath, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select input presentation', self.input_path.text(), 'PowerPoint Presentations (*.pptx)')
        if filepath:
            self.input_path.setText(str(pathlib.Path(filepath).resolve()))

    def on_size(self):
        ppt = office.PowerPoint(self.input_path.text())
        self.progress.setMaximum(ppt.doc.Slides.Count)
        self.table.setRowCount(ppt.doc.Slides.Count)
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as f:
            filepath = pathlib.Path(f.name)
        sizes = []
        for i in range(ppt.doc.Slides.Count):
            self.progress.setValue(i)
            QtCore.QCoreApplication.processEvents(QtCore.QEventLoop.AllEvents, 100)
            ppt.export(filepath, i + 1)
            sizes.append(filepath.stat().st_size / 1e6)
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(f'{sizes[i]:.2f}'))
        filepath.unlink()
        self.progress.setValue(0)
        for i, size in enumerate(sizes):
            self.table.setItem(i, 1, QtWidgets.QTableWidgetItem(f'{100 * size / sum(sizes):.0f}'))
        ppt.close(alert=False)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(description='Reports PowerPoint slide sizes')
        parser.add_argument('input', help='Input presentation')
        args = parser.parse_args()

        ppt = office.PowerPoint(args.input)
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as f:
            filepath = pathlib.Path(f.name)
        for i in range(1, ppt.doc.Slides.Count + 1):
            ppt.export(filepath, i)
            print(f'{i:>3}/{ppt.doc.Slides.Count}: {filepath.stat().st_size / 1e6:.1f} MB')
        filepath.unlink()
        del ppt
    else:
        app = QtWidgets.QApplication(sys.argv)
        window = Window()
        sys.exit(app.exec_())
