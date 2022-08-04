from win32com.client import constants


def boolean(value):
    return constants.msoTrue if value else constants.msoFalse


def inch(value, reverse=False):
    if reverse:
        return value / 72
    return value * 72


def rgb(r, g, b):
    return r + (g * 256) + (b * 256 ** 2)
