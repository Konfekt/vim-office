" Convert binary documents to text when they are read into a buffer.
" Called by the autocommands in ftdetect/office.vim.

let s:keepcpo = &cpo
set cpo&vim

scriptencoding utf-8

let s:slash = exists('+shellslash') && !&shellslash ? '\' : '/'
let s:cat   = has('win32') ? 'type' : 'cat'
let s:nul   = has('win32') ? 'NUL' : '/dev/null'

" Text browser that dumps HTML read from standard input as plain text.
let s:browser = ''
for [s:exe, s:cmd] in [
      \ ['w3m',       'w3m -dump -T text/html -s'],
      \ ['elinks',    'elinks -dump 1'],
      \ ['lynx',      'lynx -dump -stdin -trim_blank_lines'],
      \ ['html2text', 'html2text -from_encoding utf-8'],
      \ ]
  if executable(s:exe)
    let s:browser = s:cmd
    break
  endif
endfor
unlet! s:exe s:cmd

" {{{ Helpers

function! s:has_browser() abort
  return !empty(s:browser)
endfunction

" Escape a path for a filter command, including the characters (such as % and #) that ':!' expands.
function! s:shellescape(path) abort
  return shellescape(a:path, 1)
endfunction

" The file read into the current buffer, escaped for a filter command.
function! s:file() abort
  return s:shellescape(expand('%:p'))
endfunction

" Replace the buffer by the output of the shell command a:cmd.
function! s:filter(cmd) abort
  silent execute '%!' . a:cmd
endfunction

" Command that prints the HTML file a:path as plain text.
function! s:browse(path) abort
  return s:browser . ' < ' . s:shellescape(a:path)
endfunction

" Command that prints the file a:path.
function! s:cat(path) abort
  return s:cat . ' ' . s:shellescape(a:path)
endfunction

" Run a:fn with a fresh temporary directory and the path of a file named after the current one with extension a:ext in it.
" The directory is removed afterwards.
function! s:with_tempdir(ext, fn) abort
  let dir = tempname()
  call mkdir(dir, 'p')
  try
    call call(a:fn, [dir, dir . s:slash . expand('%:t:r') . '.' . a:ext])
  finally
    call delete(dir, 'rf')
  endtry
endfunction

function! s:tika(type) abort
  let cmd = 'tika --encoding=UTF-8 --detect --' . a:type . ' ' . s:file()
  call s:filter(a:type ==# 'html' ? cmd . ' | ' . s:browser : cmd)
  setlocal fileencoding=utf-8
endfunction

" Convert with LibreOffice into format a:fmt (html is shown through the text browser).
function! s:soffice(fmt) abort
  call s:with_tempdir(a:fmt, function('s:soffice_in', [a:fmt]))
endfunction

function! s:soffice_in(fmt, dir, out) abort
  call s:filter('soffice --headless --convert-to ' . a:fmt . ' --outdir ' . s:shellescape(a:dir) .
        \ ' ' . s:file() . ' > ' . s:nul . ' 2>&1 && ' .
        \ (a:fmt ==# 'html' ? s:browse(a:out) : s:cat(a:out)))
endfunction

function! s:easyalign(...) abort
  if exists(':EasyAlign') == 2
    for args in a:000
      execute '%EasyAlign' . args
    endfor
  endif
endfunction

" Show tab stops of tab-separated columns.
function! s:show_tabs() abort
  setlocal tabstop=8 nowrap list listchars=eol:\ ,tab:»-
endfunction

" Without a converter, let the zip plug-in show the archive a:exts (comma-separated patterns) instead.
function! s:zip(exts) abort
  let g:zipPlugin_ext = get(g:, 'zipPlugin_ext', '')
  let known = split(g:zipPlugin_ext, ',')
  let added = 0
  for pat in split(a:exts, ',')
    if index(known, pat) < 0
      let g:zipPlugin_ext .= (empty(g:zipPlugin_ext) ? '' : ',') . pat
    endif
    if exists('#zip') && !exists('#zip#BufReadCmd#' . pat)
      execute 'autocmd zip BufReadCmd ' . pat . ' call zip#Browse(expand("<amatch>"))'
      let added = 1
    endif
  endfor
  if added
    doautocmd zip BufReadCmd
  endif
endfunction

" Make the buffer a read-only view, after setting its filetype to a:1 (if given)
" so that its ftplugin does not undo this.
function! s:view(...) abort
  if a:0 > 0 && !empty(a:1)
    execute 'setlocal filetype=' . a:1
  endif
  setlocal readonly buftype=nowrite
  call s:undo('setl readonly< buftype<' . (a:0 > 1 ? ' ' . a:2 : ''))
endfunction

function! s:undo(cmd) abort
  let b:undo_ftplugin = (empty(get(b:, 'undo_ftplugin', '')) ? '' : b:undo_ftplugin . '|') . a:cmd
endfunction

" }}}

" {{{ PDF
function! office#pdf() abort
  if executable('pdftotext')
    call s:filter('pdftotext -enc UTF-8 -nopgbrk -q -- ' . s:file() . ' -')
    setlocal fileencoding=utf-8
  elseif executable('pdftohtml') && s:has_browser()
    call s:filter('pdftohtml -i -noframes -nodrm -enc UTF-8 -stdout ' . s:file() . ' | ' . s:browser)
    setlocal fileencoding=utf-8
  elseif executable('markitdown')
    call s:filter('markitdown ' . s:file())
  elseif executable('tika') && s:has_browser()
    call s:tika('html')
  elseif executable('tika')
    call s:tika('text')
  endif
  call s:view('text')
endfunction
" }}}

" {{{ DJVU
function! office#djvu() abort
  if executable('djvutxt')
    call s:filter('djvutxt ' . s:file() . ' -')
    setlocal fileencoding=utf-8
  endif
  call s:view('text')
endfunction
" }}}

" {{{ EPUB
function! office#epub() abort
  let ft = 'text'
  if executable('pandoc')
    call s:filter('pandoc --from=epub --to=plain ' . s:file())
    setlocal fileencoding=utf-8
  elseif executable('tika')
    call s:tika('text')
  else
    call s:zip('*.epub')
    let ft = ''
  endif
  call s:view(ft)
endfunction
" }}}

" {{{ RTF
function! office#rtf() abort
  if executable('unrtf')
    call s:filter('unrtf --text ' . s:file())
  elseif executable('tika')
    call s:tika('text')
  elseif executable('soffice')
    call s:filter('soffice --cat ' . s:file())
  endif
  call s:view('text')
endfunction
" }}}

" {{{ ODT
function! office#odt() abort
  if executable('odt2txt')
    call s:filter('odt2txt --encoding=UTF-8 ' . s:file())
    setlocal fileencoding=utf-8
  elseif executable('odt2html') && s:has_browser()
    call s:filter('odt2html ' . s:file() . ' | ' . s:browser)
    setlocal fileencoding=utf-8
  elseif executable('pandoc')
    call s:filter('pandoc --from=odt --to=plain ' . s:file())
    setlocal fileencoding=utf-8
  elseif executable('tika')
    call s:tika('text')
  elseif executable('soffice')
    call s:filter('soffice --cat ' . s:file())
  endif
  call s:view('text')
endfunction
" }}}

" {{{ ODS
function! office#ods() abort
  if executable('soffice') && s:has_browser()
    call s:soffice('html')
  elseif executable('soffice')
    call s:soffice('csv')
    call s:easyalign('*,')
    setlocal filetype=csv
  elseif executable('tika') && s:has_browser()
    call s:tika('html')
  elseif executable('tika')
    call s:tika('text')
    call s:easyalign('*/\t\+/')
    call s:show_tabs()
  elseif executable('odt2txt')
    call s:filter('odt2txt --encoding=UTF-8 ' . s:file())
    setlocal fileencoding=utf-8
  endif
  setlocal nowrap
  call s:view('', 'wrap<')
endfunction
" }}}

" {{{ ODP
function! office#odp() abort
  if executable('soffice') && s:has_browser()
    call s:soffice('html')
  elseif executable('soffice')
    call s:filter('soffice --cat ' . s:file())
    setlocal filetype=text
  elseif executable('tika') && s:has_browser()
    call s:tika('html')
    setlocal filetype=text
  elseif executable('tika')
    call s:tika('text')
    setlocal filetype=text
  elseif executable('odt2txt')
    call s:filter('odt2txt --encoding=UTF-8 ' . s:file())
    setlocal fileencoding=utf-8
    setlocal filetype=text
  endif
  call s:view()
endfunction
" }}}

" {{{ DOC
function! s:wvtext(dir, out) abort
  call s:filter('wvText ' . s:file() . ' ' . s:shellescape(a:out) . ' && ' . s:cat(a:out))
endfunction

function! s:wvhtml(dir, out) abort
  call s:filter('wvHtml --targetdir=' . s:shellescape(a:dir) . ' ' . s:file() . ' ' .
        \ s:shellescape(fnamemodify(a:out, ':t')) . ' > ' . s:nul . ' && ' . s:browse(a:out))
endfunction

function! office#doc() abort
  if executable('wvText')
    call s:with_tempdir('txt', function('s:wvtext'))
  elseif executable('wvHtml') && s:has_browser()
    call s:with_tempdir('html', function('s:wvhtml'))
  elseif executable('tika')
    call s:tika('text')
  elseif executable('soffice')
    call s:filter('soffice --cat ' . s:file())
  elseif executable('catdoc')
    call s:filter('catdoc ' . s:file())
  elseif executable('antiword')
    call s:filter('antiword ' . s:file())
  endif
  call s:view('text')
endfunction
" }}}

" {{{ DOCX
function! office#docx() abort
  if executable('doxx')
    call s:filter('doxx --export markdown ' . s:file())
    setlocal filetype=markdown
  elseif executable('docx2txt')
    " python version
    call s:filter('docx2txt ' . s:file())
    setlocal filetype=text
  elseif executable('docx2txt.pl')
    call s:filter('docx2txt.pl ' . s:file() . ' -')
    setlocal filetype=text
  elseif executable('pandoc')
    call s:filter('pandoc --from=docx --to=markdown ' . s:file())
    setlocal fileencoding=utf-8
    setlocal filetype=markdown
  elseif executable('markitdown')
    call s:filter('markitdown ' . s:file())
  elseif executable('tika')
    call s:tika('text')
    setlocal filetype=text
  elseif executable('soffice')
    call s:filter('soffice --cat ' . s:file())
    setlocal filetype=text
  else
    call s:zip('*.docx,*.docm,*.dotx,*.dotm')
  endif
  call s:view()
endfunction
" }}}

" {{{ XLS
function! s:xls2csv(dir, out) abort
  call s:filter('xls2csv -q -a UTF-8 -b WINDOWS-1252 -x ' . s:file() . ' -c ' . s:shellescape(a:out) .
        \ ' > ' . s:nul . ' && ' . s:cat(a:out))
endfunction

function! office#xls() abort
  if executable('xlhtml') && s:has_browser()
    call s:filter('xlhtml ' . s:file() . ' | ' . s:browser)
  elseif executable('xlscat')
    call s:filter('xlscat ' . s:file())
    " take account of unicode homoglyph │ instead of |
    call s:easyalign('*|', '*/│\+/')
  elseif executable('xls2csv')
    " Perl version of xls2csv = Wrapper for xlscat
    call s:with_tempdir('csv', function('s:xls2csv'))
    call s:easyalign('*,')
    setlocal filetype=csv
  elseif executable('in2csv')
    call s:filter('in2csv -f xls ' . s:file())
    call s:easyalign('*,')
  elseif executable('excel2csv')
    call s:filter('excel2csv --trim ' . s:file())
    call s:easyalign('*,')
    setlocal filetype=csv
  elseif executable('soffice') && s:has_browser()
    call s:soffice('html')
  elseif executable('soffice')
    call s:soffice('csv')
    call s:easyalign('*,')
    setlocal filetype=csv
  elseif executable('tika') && s:has_browser()
    call s:tika('html')
  elseif executable('tika')
    call s:tika('text')
    call s:easyalign('*/\t\+/')
    call s:show_tabs()
  endif
  setlocal nowrap
  call s:view('', 'wrap<')
endfunction
" }}}

" {{{ XLSX
function! office#xlsx() abort
  if executable('xlsx2csv')
    call s:filter('xlsx2csv --all --delimiter x09 --outputencoding utf-8 ' . s:file())
    setlocal fileencoding=utf-8
    if has('unix') | setlocal fileformat=unix | endif
    call s:easyalign('*/\t\+/')
    call s:show_tabs()
    " mute error message
    setlocal filetype=tsv
  elseif executable('markitdown')
    call s:filter('markitdown ' . s:file())
  elseif executable('xlscat')
    call s:filter('xlscat -S all -e ' . &encoding . ' ' . s:file())
    " take account of unicode homoglyph │ instead of |
    call s:easyalign('*|', '*/│\+/')
  elseif executable('in2csv')
    call s:filter('in2csv -f xlsx ' . s:file())
    call s:easyalign('*,')
  elseif executable('excel2csv')
    call s:filter('excel2csv ' . s:file())
    call s:easyalign('*,')
    setlocal filetype=csv
  elseif executable('soffice') && s:has_browser()
    call s:soffice('html')
  elseif executable('soffice')
    call s:soffice('csv')
    call s:easyalign('*,')
    setlocal filetype=csv
  elseif executable('tika') && s:has_browser()
    call s:tika('html')
    call s:easyalign('*/\t\+/')
    call s:show_tabs()
  elseif executable('tika')
    call s:tika('text')
    call s:easyalign('*/\t\+/')
    call s:show_tabs()
    setlocal filetype=text
  else
    call s:zip('*.xlsx,*.xlsm,*.xlsb,*.xltx,*.xltm')
  endif
  setlocal nowrap
  call s:view('', 'wrap<')
endfunction
" }}}

" {{{ PPT
function! office#ppt() abort
  if executable('tika') && s:has_browser()
    call s:tika('html')
  elseif executable('tika')
    call s:tika('text')
  elseif executable('soffice') && s:has_browser()
    call s:soffice('html')
  elseif executable('soffice')
    call s:filter('soffice --cat ' . s:file())
  elseif executable('catppt')
    call s:filter('catppt ' . s:file())
  endif
  call s:view('text')
endfunction
" }}}

" {{{ PPTX
function! s:pptx2md(dir, out) abort
  call s:filter('pptx2md --disable-image --disable-wmf --disable-color --disable-escaping ' . s:file() .
        \ ' -o ' . s:shellescape(a:out) . ' 2> ' . s:nul . ' && ' . s:cat(a:out))
endfunction

function! office#pptx() abort
  if executable('pptx2md')
    call s:with_tempdir('md', function('s:pptx2md'))
    setlocal filetype=markdown foldmethod=expr foldexpr=MarkdownFold()
    call s:undo('setl fdm< fde<')
  elseif executable('markitdown')
    call s:filter('markitdown ' . s:file())
  elseif executable('tika') && s:has_browser()
    call s:tika('html')
    setlocal filetype=text
  elseif executable('tika')
    call s:tika('text')
    setlocal filetype=text
  elseif executable('soffice') && s:has_browser()
    call s:soffice('html')
  elseif executable('soffice')
    call s:filter('soffice --cat ' . s:file())
    setlocal filetype=text
  else
    call s:zip('*.pptx,*.pptm,*.ppsx,*.ppsm')
  endif
  call s:view()
endfunction
" }}}

" {{{ OTHER FILE TYPES
function! office#tika() abort
  call s:tika('text')
  call s:view('text')
endfunction

function! office#tesseract() abort
  call s:filter('tesseract ' . get(g:, 'office_tesseract', '') . ' -c debug_file=' . s:nul .
        \ ' ' . s:file() . ' - 2> ' . s:nul)
  setlocal fileencoding=utf-8
  call s:view('text')
endfunction
" }}}

let &cpo = s:keepcpo
unlet s:keepcpo

" vim: set foldmethod=marker:
