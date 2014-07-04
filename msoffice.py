import sys
from cStringIO import StringIO

from win32com.client import constants, Dispatch, makepy
from win32com.client.gencache import EnsureDispatch


class Document(object):
    '''A minimial wrapper for managing Word through the Component Object Model (COM).
    See http://msdn.microsoft.com/en-us/library/ff837519(v=office.15).aspx.
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


class Presentation(object):
    '''A minimial wrapper for managing PowerPoint through the Component Object Model (COM).
    >>> p = Presentation()
    >>> p.set_template()
    >>> slide = p.presentation.Slides.Add(p.presentation.Slides.Count + 1, constants.ppLayoutBlank)
    >>> p.presentation.SaveAs('/path/to/presentation.pptx')
    >>> del p
    See http://msdn.microsoft.com/en-us/library/ff743835(v=office.15).aspx.
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
