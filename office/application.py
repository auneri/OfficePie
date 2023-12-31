import contextlib
import io
import os
import pathlib
import sys

import pythoncom
import win32com.client
from win32com.client import constants, makepy

from .util import boolean, inch, rgb


class Application:
    """A minimial wrapper for managing Microsoft Office documents through Component Object Model (COM).

    See https://msdn.microsoft.com/en-us/library/office/jj162978.aspx.
    """

    def __init__(self, application, document, filepath=None, visible=True, version=16.0):
        self.app = None
        self._proxy(f'Microsoft Office {version:.1f} Object Library')
        self._proxy(f'Microsoft {application} {version:.1f} Object Library')
        try:
            self.app = win32com.client.gencache.EnsureDispatch(f'{application}.Application')
        except pythoncom.com_error as error:
            raise RuntimeError(f'Failed to start {application}') from error
        if application != 'PowerPoint':
            self.app.Visible = boolean(visible)
        if filepath is not None and pathlib.Path(filepath).is_file():
            self.doc = self._get_open_file(str(filepath))
            if self.doc is None:
                kwargs = {}
                if application == 'PowerPoint':
                    kwargs['WithWindow'] = boolean(visible)
                self.doc = getattr(self.app, document).Open(str(filepath), **kwargs)
        else:
            self.doc = getattr(self.app, document).Add()
            if filepath is not None:
                self.doc.SaveAs(str(filepath))

    def quit(self):  # noqa: A003
        self.app.Quit()
        self.app = self.doc = self._proxy = None

    def close(self, alert=True, switch=None):
        if switch is None:
            switch = boolean(True), boolean(False)
        display_alerts = self.app.DisplayAlerts
        self.app.DisplayAlerts = switch[0] if alert else switch[1]
        self.doc.Close()
        self.app.DisplayAlerts = display_alerts

    def _get_open_file(self, filepath):
        context = pythoncom.CreateBindCtx(0)
        for moniker in pythoncom.GetRunningObjectTable():
            if str(filepath) == os.path.abspath(moniker.GetDisplayName(context, None)):  # noqa: PL100
                return win32com.client.GetObject(str(filepath))

    def _proxy(self, name=''):
        """Ensure generation of named static COM proxy upon dispatch."""
        f = io.StringIO()
        with contextlib.redirect_stdout(f):
            sys.argv = ['', '-i', name]
            makepy.main()
        exec('\n'.join(line.replace(' >>> ', '') for line in f.getvalue().splitlines() if line.startswith(' >>> ')))


class Word(Application):
    """Microsoft Office Word.

    >>> w = Word()
    >>> for i in range(3):
    >>>     paragraph = w.doc.Paragraphs.Add(w.doc.Paragraphs(w.doc.Paragraphs.Count).Range)
    >>>     paragraph.Range.Text = f'Paragraph {w.doc.Paragraphs.Count - 1}{os.linesep}'
    >>> w.doc.SaveAs('/path/to/file.docx')
    """

    def __init__(self, *args, **kwargs):
        super().__init__('Word', 'Documents', *args, **kwargs)

    def quit(self):  # noqa: A003
        if self.app is not None and len(self.app.Documents) > 0:
            raise RuntimeError(f'Cannot quit with {len(self.app.Documents)} document(s) open')
        super().quit()

    def add_image(self, filepath):
        paragraph = self.doc.Paragraphs.Add(self.doc.Paragraphs(self.doc.Paragraphs.Count).Range)
        return self.doc.InlineShapes.AddPicture(FileName=filepath, LinkToFile=boolean(False), SaveWithDocument=boolean(True), Range=paragraph.Range)

    def add_text(self, text):
        paragraph = self.doc.Paragraphs.Add(self.doc.Paragraphs(self.doc.Paragraphs.Count).Range)
        paragraph.Range.Text = text
        return paragraph

    def close(self, alert=True):
        super().close(alert, switch=(constants.wdAlertsAll, constants.wdAlertsNone))

    def mark_revisions(self, author=None, color=None, strike_deletions=False):
        """Convert tracked changes to marked revisions."""
        unhandled_revisions = {eval('constants.wdRevision{}'.format(revision.replace(' ', ''))): revision for revision in (  # noqa: FS002
            'Cell Deletion', ' Cell Insertion', 'Cell Merge', 'Cell Split', 'Conflict', 'Conflict Delete',
            'Conflict Insert', 'Display Field', 'Moved From', 'Moved To', 'Paragraph Number', 'Paragraph Property',
            'Property', 'Reconcile', 'Replace', 'Section Property', 'Style', 'Style Definition', 'Table Property')}

        track_revisions = self.doc.TrackRevisions
        self.doc.TrackRevisions = boolean(False)
        for i, r in enumerate(self.doc.Revisions):
            if author is None or r.Author == author:
                if r.Type == constants.wdRevisionDelete:
                    if strike_deletions:
                        r.Range.Font.ColorIndex = constants.wdBlue if color is None else color
                        r.Range.Font.StrikeThrough = boolean(True)
                        r.Reject()
                    else:
                        r.Accept()
                elif r.Type == constants.wdRevisionInsert:
                    r.Range.Font.ColorIndex = constants.wdBlue if color is None else color
                    r.Accept()
                elif r.Type == constants.wdNoRevision:
                    print('Unhandled revision: No Revision', file=sys.stderr)
                elif r.Type in unhandled_revisions:
                    print(f'Unhandled revision: {unhandled_revisions[r.Type]}', file=sys.stderr)
                else:
                    print(f'Unexpected revision type: {r.Type}', file=sys.stderr)
            yield i + 1
        self.doc.TrackRevisions = track_revisions

    def maximize(self):
        self.app.WindowState = constants.wdWindowStateMaximize


class Excel(Application):
    """Microsoft Office Excel.

    >>> e = Excel()
    >>> e.doc.SaveAs('/path/to/file.xlsx')
    """

    def __init__(self, *args, **kwargs):
        super().__init__('Excel', 'Workbooks', *args, **kwargs)

    def quit(self):  # noqa: A003
        if self.app is not None and len(self.app.Workbooks) > 0:
            raise RuntimeError(f'Cannot quit with {len(self.app.Workbooks)} workbook(s) open')
        super().quit()

    def export(self, filepath):
        self.doc.ActiveSheet.ExportAsFixedFormat(0, filepath)

    def maximize(self):
        self.app.WindowState = constants.xlMaximized


class PowerPoint(Application):
    """Microsoft Office PowerPoint.

    >>> p = PowerPoint()
    >>> slide = p.add_slide()
    >>> p.add_text(f'Slide {slide.SlideNumber}', position=(0.2,0.2), slide=slide.SlideNumber)
    >>> p.doc.SaveAs('/path/to/file.pptx')
    """

    def __init__(self, *args, **kwargs):
        super().__init__('PowerPoint', 'Presentations', *args, **kwargs)

    def quit(self):  # noqa: A003
        if self.app is not None and len(self.app.Presentations) > 0:
            raise RuntimeError(f'Cannot quit with {len(self.app.Presentations)} presentation(s) open')
        super().quit()

    def add_slide(self, layout=None):
        if layout is None:
            layout = constants.ppLayoutBlank
        elif isinstance(layout, str):
            layout = eval(f'constants.ppLayout{layout}')
        slide = self.doc.Slides.Add(self.doc.Slides.Count + 1, layout)
        slide.Select()
        return slide

    def add_text(self, text, position=(0, 0), size=(0, 0), margins=(0, 0, 0, 0), fontsize=None, fontcolor=None, bold=None, wrap=None, glow=None, slide=None):
        slide = self.get_slide(slide)
        shape = slide.Shapes.AddTextbox(Orientation=constants.msoTextOrientationHorizontal, Left=inch(position[0]), Top=inch(position[1]), Width=inch(size[0]), Height=inch(size[1]))
        shape.TextFrame.TextRange.Text = text
        if margins is not None:
            shape.TextFrame.MarginLeft = margins[0]
            shape.TextFrame.MarginRight = margins[1]
            shape.TextFrame.MarginTop = margins[2]
            shape.TextFrame.MarginBottom = margins[3]
        if fontsize is not None:
            shape.TextFrame.TextRange.Font.Size = fontsize
        if fontcolor is not None:
            shape.TextFrame.TextRange.Font.Color.RGB = rgb(*fontcolor)
        if bold is not None:
            shape.TextFrame.TextRange.Font.Bold = boolean(bold)
        if wrap is not None:
            shape.TextFrame.WordWrap = boolean(wrap)
        if glow is not None:
            shape.TextFrame2.TextRange.Font.Glow.Color.RGB = rgb(*glow['color'])
            shape.TextFrame2.TextRange.Font.Glow.Radius = glow['radius']
            shape.TextFrame2.TextRange.Font.Glow.Transparency = glow['alpha']
        return shape

    def add_image(self, filepath, position=(0, 0), size=None, slide=None):
        slide = self.get_slide(slide)
        kwargs = {}
        if size is not None:
            kwargs['Width'] = inch(size[0])
            kwargs['Height'] = inch(size[1])
        return slide.Shapes.AddPicture(FileName=filepath, LinkToFile=boolean(False), SaveWithDocument=boolean(True), Left=inch(position[0]), Top=inch(position[1]), **kwargs)

    def close(self, alert=True):
        super().close(alert, switch=(constants.ppAlertsAll, constants.ppAlertsNone))

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

    def maximize(self):
        self.app.WindowState = constants.ppWindowMaximized

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
        def ungroups(shape, shapes=[]):  # noqa: B006
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
