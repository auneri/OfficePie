#!/usr/bin/env python
"""Generate PowerPoint templates.

For help in extending this template, see https://msdn.microsoft.com/en-us/VBA/VBA-PowerPoint
"""

import argparse

import office
import winreg
from office.util import boolean, inch, rgb
from win32com.client import constants


def main(version, theme):
    if theme not in ('dark', 'light'):
        raise NotImplementedError(f'{theme} theme was not recognized')

    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, f'Software\\Microsoft\\Office\\{version:.1f}\\PowerPoint\\Options')
    winreg.SetValueEx(key, 'AutomaticPictureCompressionDefault', 0, winreg.REG_DWORD, 0)
    winreg.SetValueEx(key, 'ExportBitmapResolution', 0, winreg.REG_DWORD, int(96 * 1.5))  # 1920x1080
    winreg.CloseKey(key)

    ppt = office.PowerPoint(version=version)
    ppt.doc.RemoveDocumentInformation(constants.ppRDIDocumentProperties)

    slide_height = inch(ppt.doc.PageSetup.SlideHeight, reverse=True)
    slide_width = inch(ppt.doc.PageSetup.SlideWidth, reverse=True)
    title_height = 1.2
    padding = 0.6, 0.3
    margin = 0.1, 0.1
    indent = 0.5

    # disable "Use Timings"
    ppt.doc.SlideShowSettings.AdvanceMode = constants.ppSlideShowManualAdvance

    # assign theme fonts
    ppt.doc.SlideMaster.Theme.ThemeFontScheme.MajorFont(constants.msoThemeLatin).Name = 'Cambria'  # headings
    if theme == 'dark':
        ppt.doc.SlideMaster.Theme.ThemeFontScheme.MinorFont(constants.msoThemeLatin).Name = 'Calibri'  # body
    elif theme == 'light':
        ppt.doc.SlideMaster.Theme.ThemeFontScheme.MinorFont(constants.msoThemeLatin).Name = 'Cambria'  # body

    # assign theme colors
    if theme == 'dark':
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark1).RGB = rgb(255, 255, 255)  # white
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight1).RGB = rgb(0, 0, 0)  # black
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark2).RGB = rgb(204, 204, 204)  # dirty white
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight2).RGB = rgb(51, 51, 51)  # dirty black
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent1).RGB = rgb(238, 238, 34)  # yellow
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent2).RGB = rgb(238, 136, 238)  # magenta
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent3).RGB = rgb(34, 238, 34)  # green
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent4).RGB = rgb(34, 238, 238)  # cyan
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent5).RGB = rgb(238, 136, 34)  # orange
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorAccent6).RGB = rgb(136, 34, 238)  # purple
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorHyperlink).RGB = rgb(238, 136, 238)  # magenta
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorFollowedHyperlink).RGB = rgb(238, 136, 238)  # magenta
    elif theme == 'light':
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark1).RGB = rgb(0, 0, 0)  # black
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight1).RGB = rgb(255, 255, 255)  # white
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorDark2).RGB = rgb(51, 51, 51)  # dirty black
        ppt.doc.SlideMaster.Theme.ThemeColorScheme(constants.msoThemeColorLight2).RGB = rgb(204, 204, 204)  # dirty white
    ppt.doc.SlideMaster.Background.Fill.ForeColor.ObjectThemeColor = constants.msoThemeColorLight1

    # format slide master title
    title = ppt.doc.SlideMaster.Shapes(1)
    title.Left = inch(padding[0])
    title.Top = inch(padding[1])
    title.Width = inch(slide_width - 2 * padding[0])
    title.Height = inch(title_height)
    title.TextFrame.MarginLeft = inch(margin[0])
    title.TextFrame.MarginRight = inch(margin[0])
    title.TextFrame.MarginTop = inch(margin[1])
    title.TextFrame.MarginBottom = inch(margin[1])
    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = constants.msoThemeColorAccent1
    title.TextFrame.TextRange.Font.Size = 36
    title.TextFrame.VerticalAnchor = constants.msoAnchorMiddle

    # format slide master body
    body = ppt.doc.SlideMaster.Shapes(2)
    body.Left = inch(padding[0])
    body.Top = inch(title_height + padding[1])
    body.Width = inch(slide_width - 2 * padding[0])
    body.Height = inch(slide_height - 2 * padding[1] - title_height)
    body.TextFrame.MarginLeft = inch(margin[0])
    body.TextFrame.MarginRight = inch(margin[0])
    body.TextFrame.MarginTop = inch(margin[1])
    body.TextFrame.MarginBottom = inch(margin[1])
    body.TextFrame.VerticalAnchor = constants.msoAnchorTop
    for i, paragraph in enumerate(body.TextFrame.TextRange.Paragraphs()):
        paragraph.Font.Size = 22 - (2 * i)
        paragraph.ParagraphFormat.SpaceBefore = 1.25 * paragraph.Font.Size / (i + 1)
        body.TextFrame.Ruler.Levels(i + 1).FirstMargin = inch(indent * i)
        body.TextFrame.Ruler.Levels(i + 1).LeftMargin = inch(indent * i)
    body.TextFrame.TextRange.ParagraphFormat.Bullet.Type = constants.ppBulletNone
    body.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1

    # remove unused layouts
    for layout in tuple(ppt.doc.SlideMaster.CustomLayouts):
        if layout.Name not in ('Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Blank'):
            layout.Delete()

    # add a slide with "Title and Content"
    slide = ppt.add_slide(constants.ppLayoutBlank)

    # customize text box defaults
    shape = ppt.add_text('Defaults', position=(1, 1))
    shape.TextFrame.MarginLeft = 0
    shape.TextFrame.MarginRight = 0
    shape.TextFrame.MarginTop = 0
    shape.TextFrame.MarginBottom = 0
    shape.TextFrame.TextRange.Font.Size = 20
    shape.SetShapesDefaultProperties()
    shape.Delete()

    # customize line defaults
    shape = slide.Shapes.AddLine(inch(1), inch(1), inch(2), inch(2))
    shape.Line.Weight = 2
    shape.SetShapesDefaultProperties()
    shape.Delete()

    # customize rectangle defaults
    shape = slide.Shapes.AddShape(constants.msoShapeRectangle, inch(1), inch(1), inch(2), inch(2))
    shape.Line.ForeColor.ObjectThemeColor = constants.msoThemeColorAccent1
    shape.SetShapesDefaultProperties()
    shape.Delete()

    # create a sample content
    title_text = 'Lorem Ipsum Dolor Sit Amet'
    body_text = [
        [1, 'Lorem ipsum dolor sit amet, consectetur adipiscing elit'],
        [2, 'Nam lacinia nisl et ullamcorper luctus'],
        [1, 'Nunc vel lectus et risus maximus viverra'],
        [2, 'Morbi eget nulla sagittis, finibus quam sit amet, cursus ante'],
        [3, 'Donec luctus mauris vel tortor blandit blandit'],
        [2, 'Praesent aliquet dolor ut nisl egestas gravida']]
    title = ppt.add_text(title_text, position=(1, 1))
    title.TextFrame.TextRange.Font.Color.ObjectThemeColor = constants.msoThemeColorAccent1
    title.TextFrame.TextRange.Font.Size = 36
    title.TextFrame.TextRange.Font.Bold = boolean(True)
    body = ppt.add_text('\r'.join(x for _, x in body_text), position=(1, 2))
    body.TextFrame.TextRange.Font.Color.ObjectThemeColor = constants.msoThemeColorDark1
    body.TextFrame.TextRange.Font.Size = 20
    for i, (indent, _) in enumerate(body_text, start=1):
        body.TextFrame.TextRange.Paragraphs(i).IndentLevel = indent
    pad = 0.1
    shapes = [
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch(pad), inch(pad), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad)),
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch((slide_width + pad) / 2), inch(pad), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad)),
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch(pad), inch((slide_height + pad) / 2), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad)),
        slide.Shapes.AddShape(constants.msoShapeRectangle, inch((slide_width + pad) / 2), inch((slide_height + pad) / 2), inch((slide_width - pad) / 2 - pad), inch((slide_height - pad) / 2 - pad))]
    for i, shape in enumerate(shapes, start=2):
        shape.Line.Visible = constants.msoFalse
        shape.Fill.ForeColor.ObjectThemeColor = getattr(constants, f'msoThemeColorAccent{i}')
        shape.Fill.Transparency = 0.75
        shape.ZOrder(constants.msoSendToBack)

    return ppt


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--version', default=16.0)
    parser.add_argument('--theme', default='dark')
    args = parser.parse_args()
    ppt = main(args.version, args.theme)
