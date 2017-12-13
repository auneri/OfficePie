#!/usr/bin/env python

from __future__ import absolute_import, division, print_function

import os
import sys
from contextlib import contextmanager

import pythoncom
import win32com.client
from six import string_types
from six.moves import StringIO
from win32com.client import constants, makepy


class Office(object):
    """A minimial wrapper for managing Microsoft Office documents through Component Object Model (COM).

    See https://msdn.microsoft.com/en-us/library/office/jj162978.aspx.
    """

    def __init__(self, application, document, filepath=None, visible=None, version=16.0):
        self._proxy('Microsoft Office {:.1f} Object Library'.format(version))
        self._proxy('Microsoft {} {:.1f} Object Library'.format(application, version))
        self.app = win32com.client.gencache.EnsureDispatch('{}.Application'.format(application))
        if visible is not None:
            self.app.Visible = visible
        if filepath is not None and os.path.isfile(filepath):
            self.doc = self._get_open_file(filepath)
            if self.doc is None:
                self.doc = getattr(self.app, document).Open(filepath)
        else:
            self.doc = getattr(self.app, document).Add()
            if filepath is not None:
                self.doc.SaveAs(filepath)

    def close(self, alert=True):
        display_alerts = self.app.DisplayAlerts
        self.app.DisplayAlerts = alert
        self.doc.Close()
        self.app.DisplayAlerts = display_alerts

    def _get_open_file(self, filepath):
        context = pythoncom.CreateBindCtx(0)
        for moniker in pythoncom.GetRunningObjectTable():
            if filepath == os.path.abspath(moniker.GetDisplayName(context, None)):
                return win32com.client.GetObject(filepath)

    def _proxy(self, name=''):
        """Ensure generation of named static COM proxy upon dispatch."""
        with capture() as (stdout, _):
            sys.argv = ['', '-i', name]
            makepy.main()
        exec('\n'.join(line.replace(' >>> ', '') for line in stdout.getvalue().splitlines() if line.startswith(' >>> ')))


class Word(Office):
    """Microsoft Office Word.

    >>> w = Word()
    >>> for i in range(3):
    >>>     paragraph = w.doc.Paragraphs.Add(w.doc.Paragraphs(w.doc.Paragraphs.Count).Range)
    >>>     paragraph.Range.Text = 'Paragraph {}{}'.format(w.doc.Paragraphs.Count - 1, os.linesep)
    >>> w.doc.SaveAs('/path/to/file.docx')
    """

    def __init__(self, *args, **kwargs):
        super(Word, self).__init__('Word', 'Documents', *args, **kwargs)

    def add_image(self, filepath):
        paragraph = self.doc.Paragraphs.Add(self.doc.Paragraphs(self.doc.Paragraphs.Count).Range)
        return self.doc.InlineShapes.AddPicture(FileName=filepath, LinkToFile=False, SaveWithDocument=True, Range=paragraph.Range)

    def add_text(self, text):
        paragraph = self.doc.Paragraphs.Add(self.doc.Paragraphs(self.doc.Paragraphs.Count).Range)
        paragraph.Range.Text = text
        return paragraph

    def mark_revisions(self, author=None, color=None, strike_deletions=False):
        """Convert tracked changes to marked revisions."""
        unhandled_revisions = {eval('constants.wdRevision{}'.format(revision.replace(' ', ''))): revision for revision in (
            'Cell Deletion', ' Cell Insertion', 'Cell Merge', 'Cell Split', 'Conflict', 'Conflict Delete',
            'Conflict Insert', 'Display Field', 'Moved From', 'Moved To', 'Paragraph Number', 'Paragraph Property',
            'Property', 'Reconcile', 'Replace', 'Section Property', 'Style', 'Style Definition', 'Table Property')}

        track_revisions = self.doc.TrackRevisions
        self.doc.TrackRevisions = False
        for i, r in enumerate(self.doc.Revisions):
            if author is None or r.Author == author:
                if r.Type == constants.wdRevisionDelete:
                    if strike_deletions:
                        r.Range.Font.ColorIndex = constants.wdBlue if color is None else color
                        r.Range.Font.StrikeThrough = True
                        r.Reject()
                    else:
                        r.Accept()
                elif r.Type == constants.wdRevisionInsert:
                    r.Range.Font.ColorIndex = constants.wdBlue if color is None else color
                    r.Accept()
                elif r.Type == constants.wdNoRevision:
                    print('Unhandled revision: No Revision', file=sys.stderr)
                elif r.Type in unhandled_revisions:
                    print('Unhandled revision: {}'.format(unhandled_revisions[r.Type]), file=sys.stderr)
                else:
                    print('Unexpected revision type: {}'.format(r.Type), file=sys.stderr)
            yield i + 1
        self.doc.TrackRevisions = track_revisions


class Excel(Office):
    """Microsoft Office Excel.

    >>> e = Excel()
    >>> e.doc.SaveAs('/path/to/file.xlsx')
    """

    def __init__(self, *args, **kwargs):
        super(Excel, self).__init__('Excel', 'Workbooks', *args, **kwargs)


class PowerPoint(Office):
    """Microsoft Office PowerPoint.

    >>> p = PowerPoint()
    >>> slide = p.add_slide()
    >>> p.add_text('Slide {}'.format(slide.SlideNumber), (0.2,0.2), slide=slide.SlideNumber)
    >>> p.doc.SaveAs('/path/to/file.pptx')
    """

    def __init__(self, *args, **kwargs):
        super(PowerPoint, self).__init__('PowerPoint', 'Presentations', *args, **kwargs)

    def add_slide(self, layout=None):
        if layout is None:
            layout = constants.ppLayoutBlank
        elif isinstance(layout, string_types):
            layout = eval('constants.ppLayout{}'.format(layout))
        slide = self.doc.Slides.Add(self.doc.Slides.Count + 1, layout)
        slide.Select()
        return slide

    def add_text(self, text, position, size=(0, 0), slide=None):
        slide = self.get_slide(slide)
        shape = slide.Shapes.AddTextbox(Orientation=constants.msoTextOrientationHorizontal, Left=inch(position[0]), Top=inch(position[1]), Width=inch(size[0]), Height=inch(size[1]))
        shape.TextFrame.WordWrap = False
        shape.TextFrame.TextRange.Text = text
        return shape

    def add_image(self, filepath, position, size=None, slide=None):
        slide = self.get_slide(slide)
        if size is None:
            return slide.Shapes.AddPicture(FileName=filepath, LinkToFile=False, SaveWithDocument=True, Left=inch(position[0]), Top=inch(position[1]))
        else:
            return slide.Shapes.AddPicture(FileName=filepath, LinkToFile=False, SaveWithDocument=True, Left=inch(position[0]), Top=inch(position[1]), Width=inch(size[0]), Height=inch(size[1]))

    def get_slide(self, index=None):
        if index is None:
            index = self.app.ActiveWindow.View.Slide.SlideNumber
        elif index < 0:
            index = self.doc.Slides.Count - index + 1
        slide = self.doc.Slides(index)
        slide.Select()
        return slide

    def move_shape(self, shape, x, y, reference='center', inches=True):
        references = reference.split()
        top = references[0]
        left = references[1] if len(references) > 1 else 'center'
        if inches:
            x, y = inch(x), inch(y)
        if top == 'upper':
            shape.Top = y
        elif top == 'center':
            shape.Top = (self.doc.PageSetup.SlideHeight - shape.Height) / 2 + y
        elif top == 'lower':
            shape.Top = self.doc.PageSetup.SlideHeight - shape.Height - y
        if left == 'left':
            shape.Left = x
        elif left == 'center':
            shape.Left = (self.doc.PageSetup.SlideWidth - shape.Width) / 2 + x
        elif left == 'right':
            shape.Left = self.doc.PageSetup.SlideWidth - shape.Width - x

    def export_slides(self, path):
        if not os.path.isdir(path):
            raise IOError('Target path must be a directory')
        for i in range(1, self.doc.Slides.Count + 1):
            s = self.get_slide(i)
            s.PublishSlides(path, True, True)


@contextmanager
def capture(stdout_curr=None, stderr_curr=None):
    stdout_prev, stderr_prev = sys.stdout, sys.stderr
    if stdout_curr is None:
        stdout_curr = StringIO()
    if stderr_curr is None:
        stderr_curr = StringIO()
    try:
        sys.stdout, sys.stderr = stdout_curr, stderr_curr
        yield stdout_curr, stderr_curr
    finally:
        sys.stdout, sys.stderr = stdout_prev, stderr_prev


def inch(value, reverse=False):
    if reverse:
        return value / 72
    return value * 72


def rgb(r, g, b):
    return r + (g * 256) + (b * 256 ** 2)
