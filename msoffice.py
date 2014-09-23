#!/usr/bin/env python

'''
@author Ali Uneri
@date 2014-06-20
'''

import sys
from cStringIO import StringIO

from win32com.client import constants, Dispatch, makepy
from win32com.client.gencache import EnsureDispatch


class Document(object):
    '''A minimial wrapper for managing Word through the Component Object Model (COM).
    See http://msdn.microsoft.com/en-us/library/ff837519(v=office.15).aspx.

    >>> d = Document()
    >>> for i in range(3):
    >>>     paragraph = d.doc.Paragraphs.Add(d.doc.Paragraphs(d.doc.Paragraphs.Count).Range)
    >>>     paragraph.Range.Text = 'Paragraph {}\n'.format(d.doc.Paragraphs.Count - 1)
    >>> d.doc.SaveAs('/path/to/file.docx')
    >>> del d
    '''

    def __init__(self, path=None, visible=True):
        self.application = EnsureDispatch('Word.Application')
        self.application.Visible = visible
        if path:
            self.doc = self.application.Documents.Open(path)
        else:
            self.doc = self.application.Documents.Add()

    def __del__(self):
        self.doc.Close(False)
        self.application.Quit()

    def mark_revisions(self, strike_deletions=False):
        '''Convert tracked changes to marked revisions.'''
        track_revisions = self.doc.TrackRevisions
        self.doc.TrackRevisions = False
        for i, r in enumerate(self.doc.Revisions):
            if r.Type == constants.wdRevisionDelete:
                if strike_deletions:
                    r.Range.Font.ColorIndex = constants.wdBlue
                    r.Range.Font.StrikeThrough = True
                    r.Reject()
                else:
                    r.Accept()
            elif r.Type == constants.wdRevisionInsert:
                r.Range.Font.ColorIndex = constants.wdBlue
                r.Accept()
            elif r.Type == constants.wdNoRevision:                print >> sys.stderr, 'Unhandled revision: No Revision'
            elif r.Type == constants.wdRevisionCellDeletion:      print >> sys.stderr, 'Unhandled revision: Cell Deletion'
            elif r.Type == constants.wdRevisionCellInsertion:     print >> sys.stderr, 'Unhandled revision: Cell Insertion'
            elif r.Type == constants.wdRevisionCellMerge:         print >> sys.stderr, 'Unhandled revision: Cell Merge'
            elif r.Type == constants.wdRevisionCellSplit:         print >> sys.stderr, 'Unhandled revision: Cell Split'
            elif r.Type == constants.wdRevisionConflict:          print >> sys.stderr, 'Unhandled revision: Conflict'
            elif r.Type == constants.wdRevisionConflictDelete:    print >> sys.stderr, 'Unhandled revision: Conflict Delete'
            elif r.Type == constants.wdRevisionConflictInsert:    print >> sys.stderr, 'Unhandled revision: Conflict Insert'
            elif r.Type == constants.wdRevisionDisplayField:      print >> sys.stderr, 'Unhandled revision: Display Field'
            elif r.Type == constants.wdRevisionMovedFrom:         print >> sys.stderr, 'Unhandled revision: Moved From'
            elif r.Type == constants.wdRevisionMovedTo:           print >> sys.stderr, 'Unhandled revision: Moved To'
            elif r.Type == constants.wdRevisionParagraphNumber:   print >> sys.stderr, 'Unhandled revision: Paragraph Number'
            elif r.Type == constants.wdRevisionParagraphProperty: print >> sys.stderr, 'Unhandled revision: Paragraph Property'
            elif r.Type == constants.wdRevisionProperty:          print >> sys.stderr, 'Unhandled revision: Property'
            elif r.Type == constants.wdRevisionReconcile:         print >> sys.stderr, 'Unhandled revision: Reconcile'
            elif r.Type == constants.wdRevisionReplace:           print >> sys.stderr, 'Unhandled revision: Replace'
            elif r.Type == constants.wdRevisionSectionProperty:   print >> sys.stderr, 'Unhandled revision: Section Property'
            elif r.Type == constants.wdRevisionStyle:             print >> sys.stderr, 'Unhandled revision: Style'
            elif r.Type == constants.wdRevisionStyleDefinition:   print >> sys.stderr, 'Unhandled revision: Style Definition'
            elif r.Type == constants.wdRevisionTableProperty:     print >> sys.stderr, 'Unhandled revision: Table Property'
            else:                                                 print >> sys.stderr, 'Unexpected revision: {}'.format(r.Type)
            yield i + 1
        self.doc.TrackRevisions = track_revisions


class Presentation(object):
    '''A minimial wrapper for managing PowerPoint through the Component Object Model (COM).
    See http://msdn.microsoft.com/en-us/library/ff743835(v=office.15).aspx.

    >>> p = Presentation()
    >>> p.set_template('istar')
    >>> slide = p.ppt.Slides.Add(p.ppt.Slides.Count + 1, constants.ppLayoutBlank)
    >>> p.ppt.SaveAs('/path/to/file.pptx')
    >>> del p
    '''

    def __init__(self, path=None, version=15.0):
        win32com('Microsoft Office {:.1f} Object Library'.format(version))
        win32com('Microsoft PowerPoint {:.1f} Object Library'.format(version))
        self.application = Dispatch('PowerPoint.Application')
        self.application.Visible = True
        if path:
            self.ppt = self.application.Presentations.Open(path)
        else:
            self.ppt = self.application.Presentations.Add()

    def __del__(self):
        from pythoncom import CoUninitialize
        self.ppt.Close()
        self.application.Quit()
        CoUninitialize()

    def set_template(self, template='istar'):
        if template == 'istar':
            self.ppt.PageSetup.SlideSize = constants.ppSlideSizeOnScreen
            self.ppt.SlideMaster.Background.Fill.ForeColor.RGB = rgb(0,0,0)
            title = self.ppt.SlideMaster.TextStyles(constants.ppTitleStyle).TextFrame.TextRange
            title.Font.Name = 'Garamond'
            title.Font.Bold = True
            title.Font.Color.RGB = rgb(255, 255, 0)
            body = self.ppt.SlideMaster.TextStyles(constants.ppBodyStyle).TextFrame.TextRange
            body.Font.Name = 'Arial'
            body.Font.Color.RGB = rgb(255, 255, 255)
            body.ParagraphFormat.Bullet.Type = constants.ppBulletNone


def inch(value):
    return value * 72


def rgb(r, g, b):
    return r + (g * 256) + (b * 256 ** 2)


def win32com(name=''):
    '''Ensure generation of named static COM proxy upon dispatch.'''
    stdout = sys.stdout
    sys.stdout = StringIO()
    sys.argv = ['', '-i', name]
    makepy.main()
    output = sys.stdout.getvalue()
    sys.stdout.close()
    sys.stdout = stdout
    exec('\n'.join(output.split('\n')[3:-1]).replace(' >>> ', ''))
