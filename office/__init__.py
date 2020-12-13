import os
import sys
from contextlib import contextmanager
from io import StringIO

import pythoncom
import win32com.client
from win32com.client import constants, makepy


class Office(object):
    """A minimial wrapper for managing Microsoft Office documents through Component Object Model (COM).

    See https://msdn.microsoft.com/en-us/library/office/jj162978.aspx.
    """

    def __init__(self, application, document, filepath=None, visible=None, version=16.0):
        self._proxy('Microsoft Office {:.1f} Object Library'.format(version))
        self._proxy('Microsoft {} {:.1f} Object Library'.format(application, version))
        self.app = win32com.client.gencache.EnsureDispatch('{}.Application'.format(application))
        if visible is not None and application != 'PowerPoint':
            self.app.Visible = constants.msoTrue if visible else constants.msoFalse
        if filepath is not None and os.path.isfile(filepath):
            self.doc = self._get_open_file(filepath)
            if self.doc is None:
                if visible is not None and application == 'PowerPoint':
                    self.doc = getattr(self.app, document).Open(filepath, WithWindow=constants.msoTrue if visible else constants.msoFalse)
                else:
                    self.doc = getattr(self.app, document).Open(filepath)
        else:
            self.doc = getattr(self.app, document).Add()
            if filepath is not None:
                self.doc.SaveAs(filepath)

    def close(self, alert=True, switch=None):
        if switch is None:
            switch = constants.msoTrue, constants.msoFalse
        display_alerts = self.app.DisplayAlerts
        self.app.DisplayAlerts = switch[0] if alert else switch[1]
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

    def __del__(self):
        if len(self.app.Documents) == 0:
            self.app.Quit()

    def add_image(self, filepath):
        paragraph = self.doc.Paragraphs.Add(self.doc.Paragraphs(self.doc.Paragraphs.Count).Range)
        return self.doc.InlineShapes.AddPicture(FileName=filepath, LinkToFile=constants.msoFalse, SaveWithDocument=constants.msoTrue, Range=paragraph.Range)

    def add_text(self, text):
        paragraph = self.doc.Paragraphs.Add(self.doc.Paragraphs(self.doc.Paragraphs.Count).Range)
        paragraph.Range.Text = text
        return paragraph

    def close(self, alert=True):
        super(Word, self).close(alert, switch=(constants.wdAlertsAll, constants.wdAlertsNone))

    def mark_revisions(self, author=None, color=None, strike_deletions=False):
        """Convert tracked changes to marked revisions."""
        unhandled_revisions = {eval('constants.wdRevision{}'.format(revision.replace(' ', ''))): revision for revision in (
            'Cell Deletion', ' Cell Insertion', 'Cell Merge', 'Cell Split', 'Conflict', 'Conflict Delete',
            'Conflict Insert', 'Display Field', 'Moved From', 'Moved To', 'Paragraph Number', 'Paragraph Property',
            'Property', 'Reconcile', 'Replace', 'Section Property', 'Style', 'Style Definition', 'Table Property')}

        track_revisions = self.doc.TrackRevisions
        self.doc.TrackRevisions = constants.msoFalse
        for i, r in enumerate(self.doc.Revisions):
            if author is None or r.Author == author:
                if r.Type == constants.wdRevisionDelete:
                    if strike_deletions:
                        r.Range.Font.ColorIndex = constants.wdBlue if color is None else color
                        r.Range.Font.StrikeThrough = constants.msoTrue
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

    def __del__(self):
        if len(self.app.Workbooks) == 0:
            self.app.Quit()


class PowerPoint(Office):
    """Microsoft Office PowerPoint.

    >>> p = PowerPoint()
    >>> slide = p.add_slide()
    >>> p.add_text('Slide {}'.format(slide.SlideNumber), (0.2,0.2), slide=slide.SlideNumber)
    >>> p.doc.SaveAs('/path/to/file.pptx')
    """

    def __init__(self, *args, **kwargs):
        super(PowerPoint, self).__init__('PowerPoint', 'Presentations', *args, **kwargs)

    def __del__(self):
        if len(self.app.Presentations) == 0:
            self.app.Quit()

    def add_slide(self, layout=None):
        if layout is None:
            layout = constants.ppLayoutBlank
        elif isinstance(layout, str):
            layout = eval('constants.ppLayout{}'.format(layout))
        slide = self.doc.Slides.Add(self.doc.Slides.Count + 1, layout)
        slide.Select()
        return slide

    def add_text(self, text, position=(0, 0), size=(0, 0), slide=None):
        slide = self.get_slide(slide)
        shape = slide.Shapes.AddTextbox(Orientation=constants.msoTextOrientationHorizontal, Left=inch(position[0]), Top=inch(position[1]), Width=inch(size[0]), Height=inch(size[1]))
        shape.TextFrame.WordWrap = constants.msoFalse
        shape.TextFrame.TextRange.Text = text
        return shape

    def add_image(self, filepath, position=(0, 0), size=None, slide=None):
        slide = self.get_slide(slide)
        kwargs = {}
        if size is not None:
            kwargs['Width'] = inch(size[0])
            kwargs['Height'] = inch(size[1])
        return slide.Shapes.AddPicture(FileName=filepath, LinkToFile=constants.msoFalse, SaveWithDocument=constants.msoTrue, Left=inch(position[0]), Top=inch(position[1]), **kwargs)

    def close(self, alert=True):
        super(PowerPoint, self).close(alert, switch=(constants.ppAlertsAll, constants.ppAlertsNone))

    def export(self, filepath, index):
        self.doc.SaveCopyAs(filepath)
        other = PowerPoint(filepath, visible=False)
        for i in range(other.doc.Slides.Count, index, -1):
            other.doc.Slides(i).Delete()
        for i in range(index - 1, 0, -1):
            other.doc.Slides(i).Delete()
        other.doc.Save()
        other.close(alert=False)

    def get_slide(self, index=None):
        if index == 'master':
            return self.doc.SlideMaster
        if index is None:
            index = self.app.ActiveWindow.View.Slide.SlideNumber
        elif index >= 0:
            index += 1
        else:
            index += self.doc.Slides.Count + 1
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

    def ungroup(self, shape, flatten=False):
        def ungroups(shape, shapes=[]):
            try:
                srange = shape.Ungroup()
                for i in range(srange.Count):
                    ungroups(srange.Item(i + 1), shapes)
            except pythoncom.com_error:
                shapes.append(shape)
                return
        slide = shape.Parent
        if flatten:
            shapes = []
            ungroups(shape, shapes)
        else:
            srange = shape.Ungroup()
            shapes = [srange.Item(x + 1) for x in range(srange.Count)]
        return slide.Shapes.Range([x.Name for x in shapes]).Group() if len(shapes) > 1 else shapes[0]


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
