#!/usr/bin/env python

'''
To create a portable application, run
pyinstaller --clean --name=istarPresentationTemplate --onefile --icon=istarPresentationTemplate.ico istarPresentationTemplate.py

For help in extending this template, see
https://msdn.microsoft.com/EN-US/library/office/ee861525.aspx
'''

from win32com.client import constants
from PythonTools.helpers.Office import PowerPoint, inch, rgb

__author__ = 'Ali Uneri'


def main():
    p = PowerPoint()

    # use widescreeen format
    p.doc.PageSetup.SlideSize = constants.ppSlideSizeOnScreen16x9

    # remove unused layouts
    for layout in tuple(p.doc.SlideMaster.CustomLayouts):
        if layout.Name not in ['Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Blank']:
            layout.Delete()

    # assign theme colors
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark1).RGB = rgb(255, 255, 255)    # white
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight1).RGB = rgb(0, 0, 0)         # black
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark2).RGB = rgb(255, 255, 255)    # white
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight2).RGB = rgb(0, 0, 0)         # black
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent1).RGB = rgb(238, 238, 34)   # yellow
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent2).RGB = rgb(34, 238, 34)    # green
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent3).RGB = rgb(238, 136, 238)  # magenta
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent4).RGB = rgb(34, 238, 238)   # cyan
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent5).RGB = rgb(255, 255, 255)  # white
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent6).RGB = rgb(255, 255, 255)  # white
    p.doc.SlideMaster.Background.Fill.ForeColor.ObjectThemeColor = constants.msoThemeColorLight1

    # format slide master title
    title = p.doc.SlideMaster.Shapes(1)
    title.Left = inch(0.2)
    title.Top = inch(0.2)
    title.Width = inch(9.6)
    title.Height = inch(1)
    title.TextFrame.MarginLeft = inch(0.05)
    title.TextFrame.MarginRight = inch(0.05)
    title.TextFrame.MarginTop = inch(0.05)
    title.TextFrame.MarginBottom = inch(0.05)
    title.TextFrame.TextRange.Font.Name = 'Garamond'
    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = constants.msoThemeColorAccent1
    title.TextFrame.TextRange.Font.Size = 28
    title.TextFrame.TextRange.Font.Bold = True
    title.TextFrame.VerticalAnchor = constants.msoAnchorTop

    # format slide master body
    body = p.doc.SlideMaster.Shapes(2)
    body.Left = inch(0.2)
    body.Top = inch(1.2)
    body.Width = inch(9.6)
    body.Height = inch(4.23)
    body.TextFrame.MarginLeft = inch(0.05)
    body.TextFrame.MarginRight = inch(0.05)
    body.TextFrame.MarginTop = inch(0.05)
    body.TextFrame.MarginBottom = inch(0.05)
    body.TextFrame.TextRange.Font.Name = 'Arial'
    for i in xrange(5):
        size = 18 - (2 * i)
        body.TextFrame.TextRange.Paragraphs(i + 1).Font.Size = size
        body.TextFrame.TextRange.Paragraphs(i + 1).ParagraphFormat.SpaceBefore = size / (i + 1)
    body.TextFrame.TextRange.ParagraphFormat.Bullet.Type = constants.ppBulletNone

    # add a slide with "Title and Content"
    p.add_slide(constants.ppLayoutObject)

    # customize text box defaults
    textbox = p.add_text('', (0,0))
    textbox.TextFrame.MarginLeft = inch(0.05)
    textbox.TextFrame.MarginRight = inch(0.05)
    textbox.TextFrame.MarginTop = inch(0.05)
    textbox.TextFrame.MarginBottom = inch(0.05)
    textbox.TextFrame.TextRange.Font.Name = 'Arial'
    textbox.TextFrame.TextRange.Font.Size = 12
    textbox.SetShapesDefaultProperties()
    textbox.Delete()

    return p


if __name__ == '__main__':
    main()
