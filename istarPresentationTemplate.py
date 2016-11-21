#!/usr/bin/env python

# TODO Toggle "Do not compress images in file".
# TODO Create analogous Word template.

"""
PowerPoint template generator for I-STAR presentations.

To create a portable application, run:
    pyinstaller --clean --name=istarPresentationTemplate --onefile --windowed --icon=istarPresentationTemplate.ico istarPresentationTemplate.py

For help in extending this template, see
https://msdn.microsoft.com/EN-US/library/office/ee861525.aspx
"""

from __future__ import absolute_import, division, print_function

from win32com.client import constants

from PythonTools.helpers.Office import PowerPoint, inch, rgb

__author__ = 'Ali Uneri'


def main():
    p = PowerPoint()

    # use widescreeen format
    p.doc.PageSetup.SlideSize = constants.ppSlideSizeOnScreen16x9
    slide_height = inch(p.doc.PageSetup.SlideHeight, reverse=True)
    slide_width = inch(p.doc.PageSetup.SlideWidth, reverse=True)
    title_height = 1.0
    padding = 0.4
    margin = 0.05
    indent = 0.4

    # disable "Use Timings"
    p.doc.SlideShowSettings.AdvanceMode = constants.ppSlideShowManualAdvance

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
    title.Left = inch(padding)
    title.Top = inch(padding)
    title.Width = inch(slide_width - 2 * padding)
    title.Height = inch(title_height)
    title.TextFrame.MarginLeft = inch(margin)
    title.TextFrame.MarginRight = inch(margin)
    title.TextFrame.MarginTop = inch(margin)
    title.TextFrame.MarginBottom = inch(margin)
    title.TextFrame.TextRange.Font.Name = 'Garamond'
    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = constants.msoThemeColorAccent1
    title.TextFrame.TextRange.Font.Size = 28
    title.TextFrame.TextRange.Font.Bold = True
    title.TextFrame.VerticalAnchor = constants.msoAnchorTop

    # format slide master body
    body = p.doc.SlideMaster.Shapes(2)
    body.Left = inch(padding)
    body.Top = inch(title_height + padding)
    body.Width = inch(slide_width - 2 * padding)
    body.Height = inch(slide_height - 2 * padding - title_height)
    body.TextFrame.MarginLeft = inch(margin)
    body.TextFrame.MarginRight = inch(margin)
    body.TextFrame.MarginTop = inch(margin)
    body.TextFrame.MarginBottom = inch(margin)
    body.TextFrame.TextRange.Font.Name = 'Arial'
    for i, paragraph in enumerate(body.TextFrame.TextRange.Paragraphs()):
        paragraph.Font.Size = 18 - (2 * i)
        paragraph.ParagraphFormat.SpaceBefore = paragraph.Font.Size / (i + 1)
        body.TextFrame.Ruler.Levels(i + 1).FirstMargin = inch(indent * i)
        body.TextFrame.Ruler.Levels(i + 1).LeftMargin = inch(indent / 2.0 + indent * i)
    body.TextFrame.TextRange.ParagraphFormat.Bullet.Type = constants.ppBulletNone
    body.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.0

    # remove unused layouts
    for layout in tuple(p.doc.SlideMaster.CustomLayouts):
        if layout.Name not in ['Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Blank']:
            layout.Delete()

    # add a slide with "Title and Content"
    p.add_slide(constants.ppLayoutObject)

    # customize text box defaults
    textbox = p.add_text('', (0,0))
    textbox.TextFrame.MarginLeft = inch(margin)
    textbox.TextFrame.MarginRight = inch(margin)
    textbox.TextFrame.MarginTop = inch(margin)
    textbox.TextFrame.MarginBottom = inch(margin)
    textbox.TextFrame.TextRange.Font.Name = 'Arial'
    textbox.TextFrame.TextRange.Font.Size = 12
    textbox.SetShapesDefaultProperties()
    textbox.Delete()

    return p


if __name__ == '__main__':
    main()
