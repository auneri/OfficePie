#!/usr/bin/env python

'''
@author Ali Uneri

@todo(auneri1) Set title and text font sizes.
@todo(auneri1) Add spacing (relative to font size) before first level indentation.
@todo(auneri1) Cleanup based on initialization routine.
'''

import os
import sys
from cStringIO import StringIO
from pythoncom import CreateBindCtx, GetRunningObjectTable
from win32com.client import constants, GetObject, makepy
from win32com.client.gencache import EnsureDispatch


class Office(object):
    '''A minimial wrapper for managing Microsoft Office documents through Component Object Model (COM).
    See https://msdn.microsoft.com/en-us/library/dn833103.aspx.
    '''

    def __init__(self, application, document, filepath=None, visible=None, version=15.0):
        self._proxy('Microsoft Office {:.1f} Object Library'.format(version))
        self._proxy('Microsoft {} {:.1f} Object Library'.format(application, version))
        self.app = EnsureDispatch('{}.Application'.format(application))
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

    def _get_open_file(self, filepath):
        context = CreateBindCtx(0)
        for moniker in GetRunningObjectTable():
            if filepath == os.path.abspath(moniker.GetDisplayName(context, None)):
                return GetObject(filepath)

    def _proxy(self, name=''):
        '''Ensure generation of named static COM proxy upon dispatch.'''
        stdout = sys.stdout
        sys.stdout = StringIO()
        sys.argv = ['', '-i', name]
        makepy.main()
        output = sys.stdout.getvalue()
        sys.stdout.close()
        sys.stdout = stdout
        exec('\n'.join(output.split('\n')[3:-1]).replace(' >>> ', ''))


class Word(Office):
    '''
    >>> w = Word()
    >>> for i in range(3):
    >>>     paragraph = w.doc.Paragraphs.Add(w.doc.Paragraphs(w.doc.Paragraphs.Count).Range)
    >>>     paragraph.Range.Text = 'Paragraph {}\n'.format(w.doc.Paragraphs.Count - 1)
    >>> w.doc.SaveAs('/path/to/file.docx')
    '''

    def __init__(self, *args, **kwargs):
        super(Word, self).__init__('Word', 'Documents', *args, **kwargs)

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


class Excel(Office):
    '''
    >>> e = Excel()
    >>> e.doc.SaveAs('/path/to/file.xlsx')
    '''

    def __init__(self, *args, **kwargs):
        super(Excel, self).__init__('Excel', 'Workbooks', *args, **kwargs)


class PowerPoint(Office):
    '''
    >>> p = PowerPoint()
    >>> p.set_template('istar')
    >>> slide = p.add_slide()
    >>> p.doc.SaveAs('/path/to/file.pptx')
    '''

    def __init__(self, *args, **kwargs):
        super(PowerPoint, self).__init__('PowerPoint', 'Presentations', *args, **kwargs)

    def set_template(self, template):
        if template == 'istar':
            self.doc.PageSetup.SlideSize = constants.ppSlideSizeOnScreen16x9
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark1).RGB = rgb(255, 255, 255)    # white
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight1).RGB = rgb(0, 0, 0)         # black
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark2).RGB = rgb(238, 238, 34)     # yellow
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight2).RGB = rgb(0, 0, 0)         # black
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent1).RGB = rgb(34, 238, 34)    # green
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent2).RGB = rgb(238, 136, 238)  # magenta
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent3).RGB = rgb(34, 238, 238)   # cyan
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent4).RGB = rgb(0, 0, 0)        # black
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent5).RGB = rgb(0, 0, 0)        # black
            self.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent6).RGB = rgb(0, 0, 0)        # black
            self.doc.SlideMaster.Background.Fill.ForeColor.ObjectThemeColor = constants.msoThemeColorLight1
            for layout in list(self.doc.SlideMaster.CustomLayouts):
                if layout.Name not in ['Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Blank']:
                    layout.Delete()
            title = self.doc.SlideMaster.TextStyles(constants.ppTitleStyle).TextFrame.TextRange
            title.Font.Name = 'Garamond'
            title.Font.Color.ObjectThemeColor = constants.msoThemeColorDark2
            title.Font.Bold = True
            body = self.doc.SlideMaster.TextStyles(constants.ppBodyStyle).TextFrame.TextRange
            body.Font.Name = 'Arial'
            body.Font.Color.ObjectThemeColor = constants.msoThemeColorDark1
            body.ParagraphFormat.Bullet.Type = constants.ppBulletNone
        else:
            raise NotImplementedError('Available templates: istar')

    def add_slide(self, type=None):
        if type is None:
            type = constants.ppLayoutBlank
        return self.doc.Slides.Add(self.doc.Slides.Count + 1, type)

    def add_text(self, text, position, size, slide=-1):
        if slide == -1:
            slide = self.doc.Slides.Count
        shapes = self.doc.Slides(slide).Shapes
        shape = shapes.AddTextbox(Orientation=0x1, Left=inch(position[0]), Top=inch(position[1]), Width=inch(size[0]), Height=inch(size[1]))
        shape.TextFrame.WordWrap = False
        shape.TextFrame.TextRange.Text = text
        shape.TextFrame.TextRange.Font.Name = 'Arial'
        shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = constants.msoThemeColorDark1
        return shape

    def add_image(self, image, position, size, slide=-1):
        if slide == -1:
            slide = self.doc.Slides.Count
        shapes = self.doc.Slides(slide).Shapes
        shape = shapes.AddPicture(FileName=image, LinkToFile=False, SaveWithDocument=True, Left=inch(position[0]), Top=inch(position[1]), Width=inch(size[0]), Height=inch(size[1]))
        return shape


def inch(value):
    return value * 72


def rgb(r, g, b):
    return r + (g * 256) + (b * 256 ** 2)
