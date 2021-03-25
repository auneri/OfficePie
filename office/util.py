
def inch(value, reverse=False):
    if reverse:
        return value / 72
    return value * 72


def rgb(r, g, b):
    return r + (g * 256) + (b * 256 ** 2)
