# coding: utf-8

# Internal zspread spreadsheet functions currently supported in formulas:

from decimal import *

def SUM(values):
    ''' SUM(rangeref)'''
    ret = Decimal(0)
    for x in values:
        ret += x
    return ret

def main():
    return 0

if __name__ == '__main__':
    sys.exit( main() )

