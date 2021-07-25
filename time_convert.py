import math
def format_sec(input_value):
    b = math.floor(input_value)
    sec = int('%.0f' % ((input_value - b)*60))
    if b > 60:
        minutes = int('%.0f' % (b / 60))
        return (f'{minutes} m {b-(minutes*60)} s and {sec} ms')
    elif b < 60:
        return (f'{b} s and {sec} ms')

