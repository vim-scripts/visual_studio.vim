'''
Companion file to visual_studio.vim
Version 1.1

Copyright (c) 2003,2004 Michael Graz
mgraz.vim@plan10.com
'''

import os, sys, re, time, pywintypes, win32com.client

vsWindowKindTaskList     = '{4A9B7E51-AA16-11D0-A8C5-00A0C921A4D2}'
vsWindowKindFindResults1 = '{0F887920-C2B6-11D2-9375-0080C747D9A0}'
vsWindowKindFindResults2 = '{0F887921-C2B6-11D2-9375-0080C747D9A0}'
vsWindowKindOutput       = '{34E76E81-EE4A-11D0-AE2E-00A0C90FFFC3}'

vsBuildStateNotStarted = 1   # Build has not yet been started.
vsBuildStateInProgress = 2   # Build is currently in progress.
vsBuildStateDone = 3         # Build has been completed

#----------------------------------------------------------------------

def dte_compile_file (fn_quickfix, qf_height, qf_errorformat):
    dte = _get_dte()
    if not dte: return
    try:
        dte.ExecuteCommand ('Build.Compile')
    except Exception, e:
        _dte_exception (e)
        _vim_raise ()
        return
    # ExecuteCommand is not synchronous so we have to wait
    while dte.Solution.SolutionBuild.BuildState == vsBuildStateInProgress:
        time.sleep (0.1)
    dte_output (fn_quickfix, 'output', qf_height, qf_errorformat)
    _vim_status ('Compile file complete')
    _vim_raise ()

#----------------------------------------------------------------------

def dte_build_solution(fn_quickfix, write_first, qf_height, qf_errorformat):
    dte = _get_dte()
    if not dte: return
    if dte.CSharpProjects.Count:
        dte.Documents.CloseAll()
    elif write_first != '0':
        _dte_set_autoload ()
    # write_first is passed in as a string
    if write_first != '0':
        _vim_command ('wall')
    _dte_raise ()
    _dte_output_activate ()
    try:
        dte.Solution.SolutionBuild.Build (1)
        # Build is not synchronous so we have to wait
        while dte.Solution.SolutionBuild.BuildState == vsBuildStateInProgress:
            time.sleep (0.1)
    except Exception, e:
        _dte_exception (e)
        _vim_raise ()
        return
    dte_output (fn_quickfix, 'output', qf_height, qf_errorformat)
    _vim_status ('Build solution complete')
    _vim_raise ()

#----------------------------------------------------------------------

def dte_task_list (fn_quickfix, qf_height, qf_errorformat):
    fp_task_list = open (fn_quickfix, 'w')
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
        except: description = '<no-description>'
        print >>fp_task_list, '%s(%s) : %s' % (filename, line, description)
    fp_task_list.close ()
    _dte_quickfix_open (qf_height, qf_errorformat)
    _vim_status ('VS Task list')

#----------------------------------------------------------------------

def dte_output (fn_output, window_kind, qf_height, qf_errorformat, notify=None):
    if window_kind == 'find_results_1':
        window_name = 'Find Results 1'
        window_id = vsWindowKindFindResults1
    elif window_kind == 'find_results_2':
        window_name = 'Find Results 2'
        window_id = vsWindowKindFindResults2
    elif window_kind == 'output':
        window_name = 'Output'
        window_id = vsWindowKindOutput
    else:
        _vim_msg ('Error: unrecognized window (%s)' % window_kind)
        return
    dte = _get_dte()
    if not dte: return
    if window_id == vsWindowKindOutput:
        owp = dte.Windows.Item(window_id).Object.OutputWindowPanes.Item ('Build')
        sel = owp.TextDocument.Selection
    else:
        sel = dte.Windows.Item(window_id).Selection
    sel.SelectAll()
    fp_output = open (fn_output, 'w')
    fp_output.write (sel.Text.replace ('\r', ''))
    fp_output.close()
    sel.Collapse()
    _dte_quickfix_open (qf_height, qf_errorformat)
    # notify is passed in as a string
    if notify and notify != '0':
        _vim_status ('VS %s' % window_name)

#----------------------------------------------------------------------

def dte_get_file (modified=None):
    dte = _get_dte()
    if not dte: return
    doc = dte.ActiveDocument
    if not doc:
        _vim_status ('No VS file!')
        return
    pt = doc.Selection.ActivePoint
    file = os.path.join (doc.Path, doc.Name)
    # modified is passed in as a string
    if modified == '0':
        action = 'edit'
    else:
        action = 'split'
    lst_cmd = [
        '%s +%d %s' % (action, pt.Line, file),
        'normal %d|' % pt.DisplayColumn,
    ]
    _vim_command (lst_cmd)

#----------------------------------------------------------------------

def dte_put_file (filename, modified, line_num, col_num):
    if not filename:
        return
    dte = _get_dte()
    if not dte: return
    # modified is passed in as a string
    if modified != '0':
        _dte_set_autoload ()
        _vim_command ('write')
    io = dte.ItemOperations
    rc = io.OpenFile (os.path.abspath (filename))
    sel = dte.ActiveDocument.Selection
    sel.MoveToLineAndOffset (line_num, col_num)
    _dte_raise ()

#----------------------------------------------------------------------

def _dte_quickfix_open (qf_height, qf_errorformat):
    if qf_height > 0:
        _vim_command ('copen %s' % qf_height)
        _vim_command (r'setlocal errorformat='+qf_errorformat)
    _vim_command ('cfile')

#----------------------------------------------------------------------

def _get_dte ():
    try:
        return win32com.client.GetActiveObject ('VisualStudio.DTE')
    except pywintypes.com_error:
        _vim_msg ('Cannot access VisualStudio. Not running?')
    return None

#----------------------------------------------------------------------

_wsh = None
def _get_wsh ():
    global _wsh
    if not _wsh:
        try:
            _wsh = win32com.client.Dispatch ('WScript.Shell')
        except pywintypes.com_error:
            _vim_msg ('Cannot access WScript.Shell')
    return _wsh

#----------------------------------------------------------------------

_pid = None
def _vim_raise ():
    if _pid:
        pid = _pid
    else:
        pid = os.getpid ()
    try:
        _get_wsh().AppActivate (pid)
    except:
        pass

#----------------------------------------------------------------------

def _dte_raise ():
    dte = _get_dte()
    if not dte: return
    try:
        dte.MainWindow.Activate ()
        _get_wsh().AppActivate (dte.MainWindow.Caption)
    except:
        pass

#----------------------------------------------------------------------

def _dte_output_activate ():
    dte = _get_dte()
    if not dte: return
    dte.Windows.Item(vsWindowKindOutput).Activate()

#----------------------------------------------------------------------

def _dte_set_autoload ():
    dte = _get_dte()
    if not dte: return
    p = dte.Properties ('Environment', 'Documents')
    p.Item('DetectFileChangesOutsideIDE').Value = 1
    p.Item('AutoloadExternalChanges').Value = 1

#----------------------------------------------------------------------

def _vim_command (lst_cmd):
    if type(lst_cmd) is not type([]):
        lst_cmd = [lst_cmd]
    try:
        import vim
        has_vim = 1
    except ImportError:
        has_vim = 0
    for cmd in lst_cmd:
        if has_vim:
            vim.command (cmd)
        else:
            if cmd.startswith ('normal'):
                print 'exe "%s"' % cmd
            else:
                # Need to turn python double backslash into single backslash
                print re.sub (r'\\\\', r'\\', cmd)

#----------------------------------------------------------------------

def _dte_exception (e):
    if isinstance (e, pywintypes.com_error):
        try:
            msg = e[2][2]
        except:
            msg = None
    else:
        msg = e
    if not msg:
        msg = 'Encountered unknown exception'
    _vim_status ('ERROR %s' % msg)

#----------------------------------------------------------------------

def _vim_status (msg):
    try:
        caption = _get_dte().MainWindow.Caption.split()[0]
    except:
        caption = None
    if caption:
        msg = msg + ' (' + caption + ')'
    _vim_msg (msg)

#----------------------------------------------------------------------

def _vim_msg (msg):
    _vim_command ('echo "%s"' % re.sub ('"', '\\"', str(msg)))

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

    if sys.argv[-1].startswith ('pid='):
        global _pid
        _pid = int (sys.argv[-1][4:])
        del sys.argv [-1]

    fcn = globals()[fcn_name]
    try:
        apply(fcn, sys.argv[2:])
    except TypeError, e:
        print 'echo "ERROR in %s: %s"' % (prog, str(e))
        return

if __name__ == '__main__': main()

