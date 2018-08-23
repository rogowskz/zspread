# coding: utf-8

"""
Usage:
	cd /cygdrive/t/Zmm/fimodel/fimodel
	python <thisfile>.py export google /cygdrive/t/google-service-account-1.json "1WHpF_XapemU90iHoIB-wxAl3FcijxLnn2GsyTrjWTZ8" "RFI" "TaxData"  
	python <thisfile>.py compile <fpath> 
	python <thisfile>.py recalculate <fpath>

Example:
	cd /cygdrive/w/fimodel
	python zspread/prod/zspread.py export google /cygdrive/w/timeline/google-service-account-1.json "1WHpF_XapemU90iHoIB-wxAl3FcijxLnn2GsyTrjWTZ8" "RFI" "TaxData"  
"""

import sys
import os
import time
import locale
import re
import json

from decimal import *
getcontext().prec = 4
DEC100 = Decimal(100)

import gspread

CWD = os.getcwd()

sys.path.append(os.path.join(CWD, 'utils/prod'))
import utils

sys.path.append(os.path.join(CWD, 'gspreadutils/prod'))
import gspreadutils

import tsort

import zsf
import zsudf

locale.setlocale(locale.LC_NUMERIC, '')

# Patterns for relative RC notation of cell and range references:
RCRELCELLADR = r'(R\[([\-]?[0-9]+)\]C\[([\-]?[0-9]+)\])'
RCRELRANGEADR = r'(' + RCRELCELLADR + r':' + RCRELCELLADR + r')'

# Patterns for A1 notation of cell and range references:
A1CELLADR = r'([A-Z]+[0-9]+)'
WSREF = r'([A-Za-z][A-Za-z0-9_]*!)?'
A1CELLREF = r'(' + WSREF + r'\[' + A1CELLADR + r'\])'
A1RANGEREF = r'(' + WSREF + r'\[(' + A1CELLADR + r':' + A1CELLADR + r')\]' + r')'

# A global reference to the spreadsheet dictionary:
DD = None

def exportToJson(ws):
    range_addr = 'A1:{}'.format(ws.get_addr_int(ws.row_count, ws.col_count))
    # Download from the cloud:
    lcells = ws.range(range_addr)
    dcells = {}
    for c in lcells:
        crow = c.row
        ccol = c.col
        adr = ws.get_addr_int(crow, ccol)
        formula = c.input_value
        dcells[adr] = {'f': formula,}
    return dcells

def translateRelativeToA1Notation(dcells):
    for k in dcells:
        cell = dcells[k]
        formula = cell['f']
        crow, ccol = utils.a1ToRC(k)
        if formula.startswith('='):
            # Find range references in relative notation and translate them to A1 notation:
            for match in re.finditer(RCRELRANGEADR, formula):
                groups = match.groups()
                if groups != ():
                    start_row_offsett = int(groups[2])
                    start_col_offsett = int(groups[3])
                    end_row_offsett = int(groups[5])
                    end_col_offsett = int(groups[6])
                    start_row = crow + start_row_offsett
                    start_col = ccol + start_col_offsett
                    end_row = crow + end_row_offsett
                    end_col = ccol + end_col_offsett
                    start_a1 = utils.rcToA1((start_row, start_col))
                    end_a1 = utils.rcToA1((end_row, end_col))
                    rangeref = '[{}:{}]'.format(start_a1, end_a1)
                    formula = formula.replace(groups[0], '{}'.format(rangeref))
            # Find cell references in relative notation and translate them to A1 notation:
            for match in re.finditer(RCRELCELLADR, formula):
                groups = match.groups()
                if groups != ():
                    row_offsett = int(groups[1])
                    col_offsett = int(groups[2])
                    row = crow + row_offsett
                    col = ccol + col_offsett
                    colref = '[{}]'.format(utils.rcToA1((row, col)))
                    formula = formula.replace(groups[0], '{}'.format(colref))
            cell['f'] = formula
    return dcells

def evaluateConstant(s):
    ret = '--unknown--'
    if s == '':
        ret = ''
    elif s.isdigit() or (s[0] in ('-', '+') and s[1:].isdigit()):
        ret = Decimal(int(s))
    elif s.endswith('%'):
        try:
            ret = Decimal(s[:-1]) / DEC100
        except InvalidOperation as e:
            # Failed to parse as a decimal number, so: must be a text string:
            ret = s
    else:
        try:
            ret = Decimal(s)
        except InvalidOperation as e:
            # Failed to parse as a decimal number, so: must be a text string:
            ret = s
    return repr(ret)

def cv(celladdr):
    # Get cell value.
    v = DD['cells'][celladdr]['v']
    ret = eval(v)
    return ret if type(ret) == type(Decimal()) else Decimal(0)

def rv(rangekey):
    # Get list of range values.
    ll = [x for x in DD['ranges'][rangekey]]
    ll = [DD['cells'][x]['v'] for x in ll]
    ll = [eval(x) for x in ll]
    ll = [x for x in ll if type(x)==type(Decimal(0)) or type(x)==type(0)]
    return ll

def recalculateFormulas(dd):
    dcells = dd['cells']
    try:
        for key in dcells: # recalculating in previously established tsort order.
            cell = dcells[key]
            if 'e' in cell:
                cell['v'] = evalFormula(cell['e']) 
    except KeyError, e:
        #TODO: Implement good exception handler here:
        print
        print 'Internal error 1: Exception thrown: KeyError', key
        print 'Exiting.'
        #sys.exit()
    return dd

class ExceptionInternal(BaseException):
    pass

def doExport():
    source_type = sys.argv[2]
    if source_type == 'google':
        doExportGoogleSheets()
    elif source_type == 'microsoft':
        doExportMicrosoftExcel()
    elif source_type in ('libreoffice', 'openoffice'):
        doExportOpenOfficeSpreadsheet()
    else:
        raise ExceptionInternal('Import source not recognized: {}'.format(source_type))

def doExportMicrosoftExcel():
    #TODO:
    raise ExceptionInternal('Function not implemented: {}'.format('doExportMicrosoftExcel'))

def doExportOpenOfficeSpreadsheet():
    #TODO:
    raise ExceptionInternal('Function not implemented: {}'.format('doExportOpenOfficeSpreadsheet'))

def doExportGoogleSheets():
    client_secret_file = sys.argv[3]
    dockey = sys.argv[4]
    wsnames = sys.argv[5:]
    dworkbook = {"workbook": dockey, "merged": {"main": wsnames[0], "data": {}}, "worksheets": {}} 
    for wsname in wsnames:
        ws = gspreadutils.openWorksheet(client_secret_file, dockey, wsname)
        dcells = exportToJson(ws) 
        dworkbook["worksheets"][wsname] = {"cells": dcells}
    fpath = '{}.{}'.format(os.path.join(os.getcwd(), dockey), 'json')
    f = open(fpath, 'w')
    writeZspreadWorkbook(dworkbook, f)
    f.close()

def writeZspreadWorkbook(dworkbook, outfile):
    outfile.write('{\n')
    outfile.write('"workbook": "{}",\n'.format(dworkbook["workbook"]) )
    outfile.write('\n' )
    #outfile.write('"merged": "{}",\n'.format(dworkbook["merged"]) )
    outfile.write('"merged": {"main": ')
    outfile.write('"{}"'.format(dworkbook["merged"]["main"]))
    outfile.write(', ')
    writeZspreadWorksheet("data", dworkbook["merged"]["data"], outfile)
    outfile.write('}')
    outfile.write('\n')
    outfile.write('\n')
    if "worksheets" in dworkbook:
        outfile.write(',')
        writeZspreadWorksheets(dworkbook["worksheets"], outfile)
    outfile.write('}\n')

def writeZspreadWorksheets(dworksheets, outfile):
    outfile.write('"worksheets": {\n' )
    first = True
    for k in dworksheets:
        outfile.write('\n' )
        if not first:
            outfile.write(',')
        writeZspreadWorksheet(k, dworksheets[k], outfile)
        first = False
    outfile.write('}\n')
    outfile.write('\n' )

def writeZspreadWorksheet(wsname, dworksheet, outfile):
    outfile.write('"{}": '.format(wsname) )
    outfile.write('{\n')
    writeZspreadCells(dworksheet, outfile)
    writeZspreadRanges(dworksheet, outfile)
    writeZspreadTsort(dworksheet, outfile)
    outfile.write('}\n')

def writeZspreadCells(dworksheet, outfile):
    if "cells" in dworksheet:
        outfile.write('"cells": {')
        dcells = dworksheet["cells"]
        first = True
        for k in dcells:
            outfile.write('\n')
            if not first:
                outfile.write(',')
            outfile.write('"{}": '.format(k))
            outfile.write(json.dumps(dcells[k], separators=(",",":")))
            first = False
        outfile.write('\n}')
        outfile.write('\n')

def writeZspreadRanges(dworksheet, outfile):
    if "ranges" in dworksheet:
        outfile.write(',')
        outfile.write('"ranges": {')
        dranges = dworksheet["ranges"]
        firstk = True
        for k in dranges:
            outfile.write('\n')
            if not firstk:
                outfile.write(',')
            outfile.write('"{}": '.format(k))
            outfile.write(json.dumps(dranges[k], separators=(",",":")))
            firstk = False
        outfile.write('\n}')
        outfile.write('\n')

def writeZspreadTsort(dworksheet, outfile):
    if "tsort" in dworksheet:
        outfile.write(',')
        outfile.write('"tsort": ')
        outfile.write(json.dumps(dworksheet["tsort"], separators=(",",":")))
        outfile.write('\n')

def buildTsort(dd):
    dd = buildCellDependencies(dd)
    # Build DAG graph:
    dcells = dd['cells']
    arcs = []
    for key in dcells:
        for dependency in dcells[key]['d']:
            if dcells[dependency]["f"].startswith("="):
                arcs.append((dependency, key))
            else:
                # No need to have cells with constant values in evaluation sort order:
                pass
    # Build topologically sorted list of graph nodes:
    ts = tsort.tsort(arcs)
    # Remove cell dependencies data that are no longer needed:
    # DEBUG: Comment thia outs, to get debug output:
    for key in dcells:
        del dcells[key]['d']
    # DEBUG: Uncomment this, to get debug output:
    #dd['DAG'] = arcs
    dd['tsort'] = ts
    return dd

def buildCellDependencies(dd):
    dcells = dd['cells'] 
    dranges = dd['ranges'] 
    for key in dcells:
        ddep = {}
        formula = dcells[key]['f']
        if formula.startswith('='):
            # Find range references in A1 notation and add all range cells to direct dependencies of this cell:
            for match in re.finditer(A1RANGEREF, formula):
                groups = match.groups()
                if groups != ():
                    rkey = groups[0]
                    for rcell in dranges[rkey]:
                        ddep[rcell] = None
            # Find cell references in A1 notation and add them to direct dependencies of this cell:
            for match in re.finditer(A1CELLREF, formula):
                groups = match.groups()
                if groups != ():
                    rcell = groups[0]
                    ddep[rcell] = None
        else:
            # A value of a cell where formula is a constant literal is dependent only on this cell's formula.
            # There are no external dependenciies:
            pass
            dcells[key]['v'] = evaluateConstant(formula)
        # Add dependecies to the cell dictionary data:
        dcells[key]['d'] = ddep.keys()
    return dd

def buildRanges(dd):
    dcells = dd['cells'] 
    dranges = {}
    for key in dcells:
        formula = dcells[key]['f']
        if formula.startswith('='):
            # Find range references in A1 notation and process them:
            for match in re.finditer(A1RANGEREF, formula):
                groups = match.groups()
                if groups != ():
                    # Build list of range cells:
                    rkey = groups[0]
                    wsref = '' if groups[1] is None else groups[1]
                    a1start = groups[-2]
                    a1end = groups[-1]
                    rstart_row, rstart_col = utils.a1ToRC(a1start)
                    rend_row, rend_col = utils.a1ToRC(a1end)
                    rows = range(rstart_row, rend_row+1)
                    cols = range(rstart_col, rend_col+1)
                    rc = [(r, [c for c in cols]) for r in rows]
                    rangecells = []
                    for t in rc:
                       for t1 in t[1]:
                           rangecells.append((t[0], t1)) 
                    # Add this list to the dictionary of ranges:
                    dranges[rkey] = ['{}[{}]'.format(wsref, utils.rcToA1(x)) for x in rangecells]
    dd['ranges'] = dranges
    return dd

def evalFormula(e):
    #print e, '-->'
    ret = repr(Decimal(eval(e)))
    #print '\t-->', ret
    return ret

QUOTED = r"'([^']*)'"
NOTQUOTED = r"([^']+)"
FLOATNUM = r'([+-]?\d+\.\d+)'
INTNUM = r"(\d+)"
TRUE = r'(true)'
FALSE = r'(false)'
FUNCALL = r'([a-z][a-zA-Z0-9]+\()'
def compileFormula(formula):
    #
    # Find calls to standard internal supported functions and normalize them:
    #TODO: A bit simplistic string matching. Improve this!:
    formula = formula.replace('sum(', 'SUM(' )
    formula = formula.replace('SUM(', 'zsf.SUM(' )
    #TODO: After all internal function names are processed, process all the remaining function calls as calls to external (user defined) functions: 
    # (everything that is not either 'cv(' or 'rv(' or 'zsf.SUM(' or (maybe more here...) is a call to user defined function ...)
    for match in re.finditer(FUNCALL, formula):
        groups = match.groups()
        if groups != ():
            txt = groups[0]
            if txt not in ( 
                    'cv(',
                    'rv(',
                    'SUM(',
                ):
                formula = formula.replace(txt, "zsudf.{}".format(txt))
    #
    # Find range references in A1 notation and replace them with call to a function that gets the list of values:
    for match in re.finditer(A1RANGEREF, formula):
        groups = match.groups()
        if groups != ():
            rkey = groups[0]
            formula = formula.replace(rkey, "rv('{}')".format(rkey))
    # Find cell references in A1 notation and replace them with call to a function that gets the value of the cell:
    for match in re.finditer(A1CELLREF, formula):
        groups = match.groups()
        if groups != ():
            cellref = groups[0]
            celladdr = groups[1]
            formula = formula.replace(cellref, "cv('{}')".format(celladdr))
    #
    # (The following two search-and-replace operations MUST be in this order! (floats first integers next)):
    #
    # Find float constants and replace them with Decimal numbers::
    for match in re.finditer(FLOATNUM, formula):
        groups = match.groups()
        if groups != ():
            num = groups[0]
            formula = formula.replace(num, "Decimal('{}')".format(num))
    #
    # Find remaining integer constants and replace them with Decimal numbers:
    #     
    #     create the list of indices of 'quoted strings' and 'non-quoted fragments':
    ss = []
    for match in re.finditer(QUOTED, formula):
        groups = match.groups()
        if groups != ():
            ss.extend((match.start(), match.end()))
    ss = [0] + ss
    ss.append(len(formula))
    #     use this list:
    ff = ''
    for i in range(len(ss)-1):
        txt = formula[ss[i]:ss[i+1]]
        if not i%2:
            # process non-quoted text fragment:
            for match in re.finditer(INTNUM, txt):
                groups = match.groups()
                if groups != ():
                    num = groups[0]
                    txt = txt.replace(num, "Decimal('{}')".format(num))
        else:
            # leave quoted string alone:
            pass
        ff += txt
    formula = ff
    #
    # Find boolean constants and replace them with string equivalents:
    for match in re.finditer(TRUE, formula):
        groups = match.groups()
        if groups != ():
            val = groups[0]
            formula = formula.replace(val, "True")
    for match in re.finditer(FALSE, formula):
        groups = match.groups()
        if groups != ():
            val = groups[0]
            formula = formula.replace(val, "False")
    #
    return formula

def compileFormulas(dd):
    dcells = dd['cells']
    for key in dcells:
        cell = dcells[key]
        formula = cell['f']
        if formula.startswith('='):
            cell['e'] = compileFormula(formula[1:])
        else:
            cell['v'] = evaluateConstant(formula)
        # Delete the original formula to save space in the output data:
        # DEBUG: comment this out to get debug output:
        #del cell['f']
    return dd

def isUsingRelativeNotation(dworksheets):
    for wsname in dworksheets:
        dworksheet = dworksheets[wsname]
        dcells = dworksheet["cells"]
        for k in dcells:
            cell = dcells[k]
            formula = cell['f']
            if formula.startswith('='):
                # Find any range references in relative notation:
                for match in re.finditer(RCRELRANGEADR, formula):
                    groups = match.groups()
                    if groups != ():
                        return True
                # Find any cell references in relative notation:
                for match in re.finditer(RCRELCELLADR, formula):
                    groups = match.groups()
                    if groups != ():
                        return True
    return False

def doCompile():
    fpath = sys.argv[2]
    dworkbook = utils.readJson(fpath)
    dworksheets = dworkbook["worksheets"]
    is_using_relative_notation = isUsingRelativeNotation(dworksheets)
    for wsname in dworksheets:
        dworksheet = dworksheets[wsname]
        if is_using_relative_notation:
            translateRelativeToA1Notation(dworksheet["cells"])
        buildRanges(dworksheet)
    mergeWorksheets(dworkbook) # merge cells and ranges
    mainws = dworkbook["merged"]["data"]
    buildTsort(mainws)
    # DEBUG: Comment next line out to get debug output:
    del dworkbook["worksheets"]
    compileFormulas(mainws)
    ofpath = fpath.replace(".json", ".c.json")
    f = open(ofpath, "w")
    writeZspreadWorkbook(dworkbook, f)
    f.close()

def mergeWorksheets(dworkbook):
    mainwsname = dworkbook["merged"]["main"]
    mainwsdata = dworkbook["merged"]["data"]
    dworksheets = dworkbook["worksheets"]
    mainwsdata["cells"] = {}
    mainwsdata["ranges"] = {}
    for wsname in dworksheets:
        decorateWorksheet(wsname, dworksheets)
        dworksheet = dworksheets[wsname]
        dd = dworksheet["cells"]
        for k in dd:
            mainwsdata["cells"][k] = dd[k]
        dd = dworksheet["ranges"]
        for k in dd:
            mainwsdata["ranges"][k] = dd[k]

def decorateWorksheet(wsname, dworksheets):
    dworksheet = dworksheets[wsname]
    # Decorate cells:
    dcells = dworksheet["cells"]
    dd = {}
    for k in dcells:
        cell = dcells[k]
        formula = cell["f"]
        if formula.startswith('='):
            # Find undecorated range references and decorate them:
            for match in re.finditer(A1RANGEREF, formula):
                groups = match.groups()
                if groups != ():
                    if groups[1] is None:
                        # undecorated yet
                        formula = formula.replace(groups[0], '{}!{}'.format(wsname, groups[0]))
            # Find undecorated cell references and decorate them:
            for match in re.finditer(A1CELLREF, formula):
                groups = match.groups()
                if groups != ():
                    if groups[1] is None:
                        # undecorated yet
                        formula = formula.replace(groups[0], '{}!{}'.format(wsname, groups[0]))
            cell['f'] = formula
        dd["{}![{}]".format(wsname, k)] = cell
    dworksheets[wsname]["cells"] = dd
    # Decorate ranges:
    dranges = dworksheet["ranges"]
    dd = {}
    for k in dranges:
        rnge = dranges[k]
        groups = re.match(A1RANGEREF, k).groups()
        if groups[1] is None:
            # undecorated yet
            dd["{}!{}".format(wsname, k)] = ['{}!{}'.format(wsname, x) for x in rnge]
        else:
            # already decorated, do nohing:
            dd[k] = rnge
    dworksheets[wsname]["ranges"] = dd

def doRecalculate():
    wbdir = sys.argv[2]
    wsname = sys.argv[3]
    fpath = '{}.json'.format(os.path.join(wbdir, wsname))
    dd = utils.readJson(fpath)
    dd = buildTsort(dd)
    dd = compileFormulas(dd)
    global DD
    DD = dd
    DD = recalculateFormulas(DD)
    dd = DD
    #utils.writeJson(dd, '_' + wsname + '.out.json', indent_=4, separators_=(',', ': '))

def main():
    command = sys.argv[1]
    if command == 'export':
            doExport()
    elif command == 'compile':
            doCompile()
    elif command == 'recalculate':
            doRecalculate()
    else:
	    print 'Unrecognized command: {}'.format( command )
	    return 1
    return 0

if __name__ == '__main__':
    sys.exit( main() )

