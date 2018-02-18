#!/usr/bin/env python

"""Generates PowerPoint templates.

For help in extending this template, see https://msdn.microsoft.com/en-us/VBA/VBA-PowerPoint
"""

from __future__ import absolute_import, division, print_function

import inspect
import os
import sys

from win32com.client import constants

sys.path.insert(0, os.path.abspath(os.path.join(inspect.getfile(inspect.currentframe()), '..', '..', '..')))
from office import inch, PowerPoint, rgb  # noqa: E402, I100, I202


def main():
    version = 16.0

    import winreg
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, 'Software\\Microsoft\\Office\\{:.1f}\\PowerPoint\\Options'.format(version))
    winreg.SetValueEx(key, 'AutomaticPictureCompressionDefault', 0, winreg.REG_DWORD, 0)
    winreg.SetValueEx(key, 'ExportBitmapResolution', 0, winreg.REG_DWORD, 144)
    winreg.CloseKey(key)

    p = PowerPoint(version=version)

    slide_height = inch(p.doc.PageSetup.SlideHeight, reverse=True)
    slide_width = inch(p.doc.PageSetup.SlideWidth, reverse=True)
    title_height = 1.2
    padding = 0.6, 0.3
    margin = 0.1, 0.1
    indent = 0.75

    # disable "Use Timings"
    p.doc.SlideShowSettings.AdvanceMode = constants.ppSlideShowManualAdvance

    # assign theme colors
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark1).RGB = rgb(255, 255, 255)    # white
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight1).RGB = rgb(0, 0, 0)         # black
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark2).RGB = rgb(204, 204, 204)    # dirty white
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight2).RGB = rgb(51, 51, 51)      # dirty black
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent1).RGB = rgb(238, 238, 35)   # yellow
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent2).RGB = rgb(238, 136, 238)  # magenta
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent3).RGB = rgb(35, 238, 35)    # green
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent4).RGB = rgb(35, 238, 238)   # cyan
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent5).RGB = rgb(238, 136, 35)   # orange
    p.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent6).RGB = rgb(136, 35, 238)   # purple
    p.doc.SlideMaster.Background.Fill.ForeColor.ObjectThemeColor = constants.msoThemeColorLight1

    # format slide master title
    title = p.doc.SlideMaster.Shapes(1)
    title.Left = inch(padding[0])
    title.Top = inch(padding[1])
    title.Width = inch(slide_width - 2 * padding[0])
    title.Height = inch(title_height)
    title.TextFrame.MarginLeft = inch(margin[0])
    title.TextFrame.MarginRight = inch(margin[0])
    title.TextFrame.MarginTop = inch(margin[1])
    title.TextFrame.MarginBottom = inch(margin[1])
    title.TextFrame.TextRange.Font.Name = 'Garamond'
    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = constants.msoThemeColorAccent1
    title.TextFrame.TextRange.Font.Size = 36
    title.TextFrame.TextRange.Font.Bold = constants.msoTrue
    title.TextFrame.VerticalAnchor = constants.msoAnchorMiddle

    # format slide master body
    body = p.doc.SlideMaster.Shapes(2)
    body.Left = inch(padding[0])
    body.Top = inch(title_height + padding[1])
    body.Width = inch(slide_width - 2 * padding[0])
    body.Height = inch(slide_height - 2 * padding[1] - title_height)
    body.TextFrame.MarginLeft = inch(margin[0])
    body.TextFrame.MarginRight = inch(margin[0])
    body.TextFrame.MarginTop = inch(margin[1])
    body.TextFrame.MarginBottom = inch(margin[1])
    body.TextFrame.TextRange.Font.Name = 'Arial'
    body.TextFrame.VerticalAnchor = constants.msoAnchorTop
    for i, paragraph in enumerate(body.TextFrame.TextRange.Paragraphs()):
        paragraph.Font.Size = 22 - (2 * i)
        paragraph.ParagraphFormat.SpaceBefore = paragraph.Font.Size / (i + 1)
        body.TextFrame.Ruler.Levels(i + 1).FirstMargin = inch(indent * i)
        body.TextFrame.Ruler.Levels(i + 1).LeftMargin = inch(indent / 2 + indent * i)
    body.TextFrame.TextRange.ParagraphFormat.Bullet.Type = constants.ppBulletNone
    body.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1

    # remove unused layouts
    for layout in tuple(p.doc.SlideMaster.CustomLayouts):
        if layout.Name not in ['Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Blank']:
            layout.Delete()

    # add a slide with "Title and Content"
    slide = p.add_slide(constants.ppLayoutObject)

    # customize text box defaults
    shape = p.add_text('', (0, 0))
    shape.TextFrame.MarginLeft = inch(margin[0])
    shape.TextFrame.MarginRight = inch(margin[0])
    shape.TextFrame.MarginTop = inch(margin[1])
    shape.TextFrame.MarginBottom = inch(margin[1])
    shape.TextFrame.TextRange.Font.Name = 'Arial'
    shape.TextFrame.TextRange.Font.Size = 20
    shape.SetShapesDefaultProperties()
    shape.Delete()

    # customize line defaults
    shape = slide.Shapes.AddLine(inch(1), inch(1), inch(2), inch(2))
    shape.Line.Weight = 1.5
    shape.SetShapesDefaultProperties()
    shape.Delete()

    slide = p.add_slide(constants.ppLayoutBlank)
    pad = 0.1
    shapes = [
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch(pad), inch(pad), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad)),
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch((slide_width + pad) / 2), inch(pad), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad)),
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch(pad), inch((slide_height + pad) / 2), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad)),
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch((slide_width + pad) / 2), inch((slide_height + pad) / 2), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad))]
    for i, shape in enumerate(shapes, start=1):
        shape.Line.Visible = constants.msoFalse
        shape.Fill.ForeColor.ObjectThemeColor = getattr(constants, 'msoThemeColorAccent{}'.format(i))
        shape.Fill.Transparency = 0.5


if __name__ == '__main__':
    main()
