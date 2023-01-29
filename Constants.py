from openpyxl.styles import Color, PatternFill, Font, Border

months = {
    1: 'May',
    2: 'June',
    3: 'July',
    4: 'August',
    5: 'September',
    6: 'October',
    7: 'November',
    8: 'December',
    9: 'January',
    10: 'February',
    11: 'March',
    12: 'April',
}

cost_centers = {
    511200: 'R&D',
    511201: 'QA',
    511202: 'Finance',
    511203: 'Admin',
    511204: 'HR',
    511205: 'BOD',
    511206: 'General Mkt',
    510686: 'R&D General',
}
alpha = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

grand_total_text = 'Grand Total'
comments_text = 'Comments'

def num_hash(num):
    if num < 26:
        return alpha[num-1]
    else:
        q, r = num//26, num % 26
        if r == 0:
            if q == 1:
                return alpha[r-1]
            else:
                return num_hash(q-1) + alpha[r-1]
        else:
            return num_hash(q) + alpha[r-1]

def get_fill(name):
    if name == 'cc':
        fill = PatternFill(patternType='solid', fgColor='f1c232')
    elif name == 'title':
        fill = PatternFill(patternType='solid', fgColor='D9E1F2')
    return fill


