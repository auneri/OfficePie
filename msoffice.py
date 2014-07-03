import sys
from cStringIO import StringIO

from pythoncom import CoUninitialize
from win32com.client import constants, Dispatch, makepy


class Presentation(object):
    '''A minimial wrapper around win32com functionalty for creating PowerPoint presentations.
    >>> p = Presentation()
    >>> slide = p.presentation.Slides.Add(p.presentation.Slides.Count + 1, constants.ppLayoutBlank)
    >>> p.presentation.SaveAs('/path/to/presentation.pptx')
    >>> del p
    '''

    def __init__(self, version=15.0, template='istar'):
        win32com('Microsoft Office {:.1f} Object Library'.format(version))
        win32com('Microsoft PowerPoint {:.1f} Object Library'.format(version))
        self.application = Dispatch('PowerPoint.Application')
        self.application.Visible = True
        self.presentation = self.application.Presentations.Add()
        self._set_template(template)

    def __del__(self):
        self.presentation.Close()
        self.application.Quit()
        CoUninitialize()

    def _set_template(self, template):
        if template == 'istar':
            self.presentation.PageSetup.SlideSize = constants.ppSlideSizeOnScreen
            self.presentation.SlideMaster.Background.Fill.ForeColor.RGB = rgb(0,0,0)
            title = self.presentation.SlideMaster.TextStyles(constants.ppTitleStyle).TextFrame.TextRange
            title.Font.Name = 'Garamond'
            title.Font.Bold = True
            title.Font.Color.RGB = rgb(255, 255, 0)
            body = self.presentation.SlideMaster.TextStyles(constants.ppBodyStyle).TextFrame.TextRange
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
