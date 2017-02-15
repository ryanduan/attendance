# -*- coding: utf-8 -*-

"""
Create at 16/11/28
"""

__author__ = 'TT'

from calendar import monthrange, weekday


y = 2016
m = 11
d = 28

print weekday(y, m, d)

d = ['2016/11/2']
print [int(day.split('/')[-1]) for day in d]
