" Visual Studio .NET integration with Vim
"
" Copyright (c) 2003,2004 Michael Graz
" mgraz.vim@plan10.com
"
" Version 1.1 May-04
" Support for compiling & building
" Thanks to Leif Wickland for contributing to this feature
"
" Version 1.0 Dec-03
" Base functionality

if exists('loaded_plugin_visual_studio')
    finish
endif
let loaded_plugin_visual_studio = 1

"----------------------------------------------------------------------
" Menu setup

if has('gui') && ( ! exists('g:visual_studio_menu') || g:visual_studio_menu != 0 )
    amenu <silent> &VisualStudio.&Get\ File :call DTEGetFile()<cr>
    amenu <silent> &VisualStudio.&Put\ File :call DTEPutFile()<cr>
    amenu <silent> &VisualStudio.&Task\ List :call DTETaskList()<cr>
    amenu <silent> &VisualStudio.&Output :call DTEOutput()<cr>
    amenu <silent> &VisualStudio.&Find\ Results\ 1 :call DTEFindResults(1)<cr>
    amenu <silent> &VisualStudio.Find\ Results\ &2 :call DTEFindResults(2)<cr>
    amenu <silent> &VisualStudio.&Compile :call DTECompileFile()<cr>
    amenu <silent> &VisualStudio.&Build\ Solution :call DTEBuildSolution()<cr>
endif

"----------------------------------------------------------------------
" Mapping setup

if ! exists ('g:visual_studio_mapping') || g:visual_studio_mapping != 0
    nmap <silent> <Leader>vp :call DTEPutFile()<cr>
    nmap <silent> <Leader>vt :call DTETaskList()<cr>
    nmap <silent> <Leader>vo :call DTEOutput()<cr>
    nmap <silent> <Leader>vf :call DTEFindResults(1)<cr>
    nmap <silent> <Leader>v2 :call DTEFindResults(2)<cr>
    nmap <silent> <Leader>vg :call DTEGetFile()<cr>
    nmap <silent> <Leader>vc :call DTECompileFile()<cr>
    nmap <silent> <Leader>vb :call DTEBuildSolution()<cr>
endif

"----------------------------------------------------------------------
" Global variables

" If setting special versions of the following vs_ files,
" make sure to escape backslashes.

if ! exists ('g:visual_studio_task_list')
    let g:visual_studio_task_list = escape($TEMP,'\').'\\vs_task_list.txt'
endif
if ! exists ('g:visual_studio_output')
    let g:visual_studio_output = escape($TEMP,'\').'\\vs_output.txt'
endif
if ! exists ('g:visual_studio_find_results_1')
    let g:visual_studio_find_results_1 = escape($TEMP,'\').'\\vs_find_results_1.txt'
endif
if ! exists ('g:visual_studio_find_results_2')
    let g:visual_studio_find_results_2 = escape($TEMP,'\').'\\vs_find_results_2.txt'
endif
if ! exists ('g:visual_studio_quickfix_height')
    let g:visual_studio_quickfix_height = 20
endif
if ! exists ('g:visual_studio_quickfix_errorformat')
    let g:visual_studio_quickfix_errorformat =
\       escape ('\ %#%f(%l)\ :\ %t%*\\D%n:\ %m,\\\ %#%f(%l\\,%c)\:\ %t%*\\D%n:\ %m', '\')
endif
if ! exists ('g:visual_studio_python_exe')
    let g:visual_studio_python_exe = 'python.exe'
endif
if ! exists ('g:visual_studio_write_before_build')
    let g:visual_studio_write_before_build = 1
endif

"----------------------------------------------------------------------
" Local variables

let s:visual_studio_module = 'visual_studio'
let s:visual_studio_python_init = 0
let s:visual_studio_location = expand("<sfile>:h")
let s:visual_studio_has_python = has('python')

"----------------------------------------------------------------------

function! <Sid>PythonInit()
    if s:visual_studio_python_init
        return 1
    endif
    if s:visual_studio_has_python
        python import sys
        exe 'python sys.path.append(r"'.s:visual_studio_location.'")'
        exe 'python import '.s:visual_studio_module
    else
        if ! <Sid>PythonCheck()
            return 0
        endif
        let s:visual_studio_module = '"'.s:visual_studio_location.'\'.s:visual_studio_module.'.py"'
        let s:visual_studio_module = escape (s:visual_studio_module, '\')
    endif
    let s:visual_studio_python_init = 1
    return 1
endfunction

"----------------------------------------------------------------------

function! <Sid>PythonCheck()
    let python_version = system (g:visual_studio_python_exe.' -V')
    if python_version !~? '^Python'
        echo 'ERROR cannot run: '.g:visual_studio_python_exe
        echo 'Update the system PATH or else set g:visual_studio_python_exe to a valid python.exe'
        return 0
    else
        return 1
    endif
endfunction

"----------------------------------------------------------------------

function! <Sid>GetPid()
    try
        let pid = libcallnr ('msvcrt.dll', '_getpid', '')
        return pid
    catch
    endtry
    try
        let pid = libcallnr ('crtdll.dll', '_getpid', '')
        return pid
    catch
    endtry
    return 0
endfunction

"----------------------------------------------------------------------

function! <Sid>DTEExec(fcn_py, ...)
    if ! <Sid>PythonInit()
        return
    endif
    " Build up args string
    let args = ''
    let i=1
    while i<= a:0
        if i > 1
            if s:visual_studio_has_python
                let args = args . ','
            else
                let args = args . ' '
            endif
        endif
        exe 'let arg=a:'.i
        let args = args.'"'.arg.'"'
        let i=i+1
    endwhile

    if s:visual_studio_has_python
        exe 'python '.s:visual_studio_module.'.'.a:fcn_py.'('.args.')'
    else
        let pid = <Sid>GetPid()
        if pid > 0
            let args = args . ' pid=' . pid
        endif
        exe system(g:visual_studio_python_exe.' '.s:visual_studio_module.' '.a:fcn_py.' '.args)
    endif
endfunction

"----------------------------------------------------------------------

function! DTEGetFile()
    call <Sid>DTEExec ('dte_get_file', &modified)
endfunction

"----------------------------------------------------------------------

function! DTEPutFile()
    let filename = escape(expand('%:p'),'\')
    if filename == ''
        echo 'No vim file!'
        return 0
    endif
"     XX TODO remove
"     if &modified == 1
"         write
"     endif
    call <Sid>DTEExec ('dte_put_file', filename, &modified, line('.'), col('.'))
    return 1
endfunction

"----------------------------------------------------------------------

function! DTETaskList()
    let &errorfile = g:visual_studio_task_list
    call <Sid>DTEExec ('dte_task_list', &errorfile,
\       g:visual_studio_quickfix_height, '%f(%l)\ %#:\ %#%m')
endfunction

"----------------------------------------------------------------------

function! DTEOutput()
    let &errorfile = g:visual_studio_output
    call <Sid>DTEExec ('dte_output', &errorfile, 'output',
\       g:visual_studio_quickfix_height, g:visual_studio_quickfix_errorformat, 1)
endfunction

"----------------------------------------------------------------------

function! DTEFindResults(which)
    if a:which == 1
        let &errorfile = g:visual_studio_find_results_1
        let window_kind = 'find_results_1'
    else
        let &errorfile = g:visual_studio_find_results_2
        let window_kind = 'find_results_2'
    endif
    call <Sid>DTEExec ('dte_output', &errorfile, window_kind,
\       g:visual_studio_quickfix_height, '%f(%l)\ %#:\ %#%m', 1)
endfunction

"----------------------------------------------------------------------

function! DTECompileFile()
    if ! DTEPutFile()
        return
    endif
    " Need to cd to the directory of the file
    " since error results have no path component
    if strlen(expand('%:h')) > 0
        cd %:h
    endif
    let &errorfile = g:visual_studio_output
    call <Sid>DTEExec ('dte_compile_file', &errorfile,
\       g:visual_studio_quickfix_height, g:visual_studio_quickfix_errorformat)
endfunction

"----------------------------------------------------------------------

function! DTEBuildSolution()
    let &errorfile = g:visual_studio_output
    call <Sid>DTEExec ('dte_build_solution', &errorfile,
\       g:visual_studio_write_before_build,
\       g:visual_studio_quickfix_height, g:visual_studio_quickfix_errorformat)
endfunction

