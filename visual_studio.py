'''
Visual Studio .NET integration with Vim
Companion file to visual_studio.vim

Copyright (c) 2003 Michael Graz
mgraz.vim@plan10.com
'''

import os, sys, pywintypes, win32com.client

vsWindowKindTaskList     = '{4A9B7E51-AA16-11D0-A8C5-00A0C921A4D2}'
vsWindowKindFindResults1 = '{0F887920-C2B6-11D2-9375-0080C747D9A0}'
vsWindowKindFindResults2 = '{0F887921-C2B6-11D2-9375-0080C747D9A0}'

#----------------------------------------------------------------------

def dte_task_list (fn_task_list):
    fp_task_list = open (fn_task_list, 'w')
    dte = _get_dte()
    if not dte: return
    TL = dte.Windows.Item(vsWindowKindTaskList).Object
    for i in range (1, TL.TaskItems.Count+1):
        TLItem = TL.TaskItems.Item(i)
        try: filename = TLItem.FileName
        except: filename = '<no-filename>'
        try: line = TLItem.Line
        except: line = '<no-line>'
        try: description = TLItem.Description
        except: Description = '<no-description>'
        print >>fp_task_list, '%s(%s) : %s' % (filename, line, description)
    fp_task_list.close()

#----------------------------------------------------------------------

def dte_output (fn_output, window_kind):
    if window_kind == 'find_results_1':
        window_id = vsWindowKindFindResults1
    elif window_kind == 'find_results_2':
        window_id = vsWindowKindFindResults2
    else:
        _vim_command ('echo "Error: unrecognized window (%s)"' % window_kind)
        return
    dte = _get_dte()
    if not dte: return
    sel = dte.Windows.Item(window_id).Selection
    sel.SelectAll()
    fp_output = open (fn_output, 'w')
    fp_output.write (sel.Text.replace ('\r', ''))
    fp_output.close()
    sel.Collapse()

#----------------------------------------------------------------------

def dte_get_file ():
    dte = _get_dte()
    if not dte: return
    doc = dte.ActiveDocument
    if not doc:
        _vim_command ('echo "No VS file!"')
        return
    pt = doc.Selection.ActivePoint
    file = os.path.join (doc.Path, doc.Name)
    lst_cmd = [
        'edit +%d %s' % (pt.Line, file),
        'normal %d|' % pt.DisplayColumn,
    ]
    _vim_command (lst_cmd)

#----------------------------------------------------------------------

def dte_put_file (filename, line_num, col_num):
    if not filename:
        return
    dte = _get_dte()
    if not dte: return
    io = dte.ItemOperations
    io.OpenFile (os.path.abspath (filename))
    sel = dte.ActiveDocument.Selection
    sel.MoveToLineAndOffset (line_num, col_num)

#----------------------------------------------------------------------

def _get_dte ():
    try:
        return win32com.client.GetActiveObject ('VisualStudio.DTE')
    except pywintypes.com_error:
        _vim_command ('echo "Cannot access VisualStudio.  Not running?"')
        return 0

#----------------------------------------------------------------------

def _vim_command (lst_cmd):
    if type(lst_cmd) is not type([]):
        lst_cmd = [lst_cmd]
    has_vim = 0
    try:
        import vim
        has_vim = 1
    except ImportError:
        has_vim = 0

    if has_vim:
        for cmd in lst_cmd:
            vim.command (cmd)
    else:
        lst_result = []
        for cmd in lst_cmd:
            if cmd.startswith ('normal'):
                lst_result.append ('exe "%s"' % cmd)
            else:
                lst_result.append (cmd)
        print '|'.join(lst_result)

#----------------------------------------------------------------------

def main ():
    prog = os.path.basename(sys.argv[0])
    if len(sys.argv) == 1:
        print 'echo "ERROR: not enough args to %s"' % prog
        return

    fcn_name = sys.argv[1]
    if not globals().has_key(fcn_name):
        print 'echo "ERROR: no such fcn %s in %s"' % (fcn_name, prog)
        return

    fcn = globals()[fcn_name]
    try:
        apply(fcn, sys.argv[2:])
    except TypeError, e:
        print 'echo "ERROR in %s: %s"' % (prog, str(e))
        return

if __name__ == '__main__': main()

