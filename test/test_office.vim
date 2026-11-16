" Tests for vim-office; run by test/run.sh, which runs each Test_* function in a fresh Vim.
" Converters are either the real tools (the test is skipped if missing) or test/stub.sh.

let s:dir = expand('<sfile>:p:h')
let s:fixtures = s:dir . '/fixtures'
let s:tmp = tempname()
call mkdir(s:tmp, 'p')
let s:errors = []
let s:path = $PATH

" Restrict $PATH to the basic utilities plus the given tools: a name links the
" stub, a name of the form 'real:NAME' links the installed tool if available.
function! s:Tools(...) abort
  let bin = s:tmp . '/bin'
  call mkdir(bin, 'p')
  for tool in ['sh', 'cat', 'sed', 'basename', 'rm', 'mkdir', 'unzip']
    let path = exepath(tool)
    if !empty(path) | call system('ln -sf ' . shellescape(path) . ' ' . shellescape(bin . '/' . tool)) | endif
  endfor
  for tool in a:000
    if tool =~# '^real:'
      let path = exepath(tool[5:])
      if empty(path) | throw 'SKIP: ' . tool[5:] . ' not installed' | endif
      " a wrapper rather than a link, so that the tool finds its own helpers
      call writefile(['#!/bin/sh', 'PATH=' . shellescape(s:path) . ' exec ' . shellescape(path) . ' "$@"'], bin . '/' . tool[5:])
      call setfperm(bin . '/' . tool[5:], 'rwxr-xr-x')
    else
      call system('ln -sf ' . shellescape(s:dir . '/stub.sh') . ' ' . shellescape(bin . '/' . tool))
    endif
  endfor
  let $PATH = bin
endfunction

function! s:Start(...) abort
  set shell=/bin/sh
  filetype plugin on
  if a:0 && a:1 ==# 'zip'
    runtime plugin/zipPlugin.vim
  endif
endfunction

" Copy fixture a:src (or create a file with content a:src if no such fixture) to a:name in the temporary directory.
function! s:File(name, src) abort
  let path = s:tmp . '/' . a:name
  if filereadable(s:fixtures . '/' . a:src)
    call writefile(readfile(s:fixtures . '/' . a:src, 'b'), path, 'b')
  else
    call writefile([a:src], path)
  endif
  return path
endfunction

function! s:Open(path) abort
  execute 'edit ' . fnameescape(a:path)
  return join(getline(1, '$'), "\n")
endfunction

function! s:Check(cond, msg) abort
  if !a:cond | call add(s:errors, a:msg) | endif
endfunction

function! s:Match(text, pat) abort
  call s:Check(a:text =~# a:pat, 'expected ' . string(a:pat) . ' in ' . string(a:text))
endfunction

function! s:ReadOnlyView(ft) abort
  call s:Check(&l:readonly, 'buffer not readonly')
  call s:Check(&l:buftype ==# 'nowrite', 'buftype is ' . string(&l:buftype))
  call s:Check(&l:filetype ==# a:ft, 'filetype is ' . string(&l:filetype))
  call s:Match(get(b:, 'undo_ftplugin', ''), '\<setl .*readonly<')
  call s:Match(get(b:, 'undo_ftplugin', ''), '\<setl .*buftype<')
endfunction

" Bug fixes {{{1

" Setting the filetype ran b:undo_ftplugin, which made the buffer writable again.
function! Test_rtf_stays_readonly() abort
  call s:Tools()
  call s:Start()
  call s:Open(s:File('a.rtf', '{\rtf1 hello}'))
  call s:ReadOnlyView('text')
endfunction

function! Test_pdf_undo_ftplugin_keeps_view_settings() abort
  call s:Tools('pdftotext')
  call s:Start()
  call s:Match(s:Open(s:File('a.pdf', 'raw')), 'STUB pdftotext')
  call s:ReadOnlyView('text')
  " the text ftplugin's own undo is kept as well
  call s:Match(b:undo_ftplugin, 'comments<.*|setl readonly< buftype<$')
endfunction

" The pattern *.djvu? did not match *.djvu.
function! Test_djvu_is_converted() abort
  call s:Tools('djvutxt')
  call s:Start()
  call s:Match(s:Open(s:File('a.djvu', 'raw')), '^STUB djvutxt .*a\.djvu -$')
  call s:ReadOnlyView('text')
endfunction

" The zip fallback ran even after a successful conversion, replacing it with
" the archive listing, and intercepted every later EPUB.
function! Test_epub_conversion_not_replaced_by_zip_listing() abort
  call s:Tools('real:pandoc')
  call s:Start('zip')
  call s:Match(s:Open(s:File('a.epub', 'hello.epub')), 'Hello from epub')
  call s:ReadOnlyView('text')
  call s:Match(s:Open(s:File('b.epub', 'hello.epub')), 'Hello from epub')
endfunction

" The odt2html branch ran odt2txt.
function! Test_odt2html_is_run() abort
  call s:Tools('odt2html', 'lynx')
  call s:Start()
  call s:Match(s:Open(s:File('a.odt', 'raw')), 'STUB odt2html .*a\.odt')
endfunction

" The wvHtml branch ran wvText without an output file.
function! Test_wvhtml_is_run() abort
  call s:Tools('wvHtml', 'lynx')
  call s:Start()
  call s:Match(s:Open(s:File('a.doc', 'raw')), '^STUB wvHtml$')
endfunction

" With $TMPDIR unset (the common case on Linux) LibreOffice got no output directory.
function! Test_soffice_without_tmpdir() abort
  call s:Tools('soffice')
  let $TMPDIR = ''
  call s:Start()
  call s:Match(s:Open(s:File('a.ods', 'raw')), '^STUB soffice csv$')
  call s:Check(&l:filetype ==# 'csv', 'filetype is ' . &l:filetype)
endfunction

" The HTML file written by LibreOffice was passed as an argument, which
" 'lynx -stdin' ignores; it now reads the file from standard input.
function! Test_soffice_html_through_lynx() abort
  call s:Tools('soffice', 'lynx')
  call s:Start()
  call s:Match(s:Open(s:File('a.ods', 'raw')), '^STUB soffice html$')
endfunction

" ':!' expanded # and % in the escaped file name.
function! Test_special_characters_in_file_name() abort
  call s:Tools('tika')
  call s:Start()
  call s:Match(s:Open(s:File('a#1 50%.ppt', 'raw')), 'STUB tika --encoding=UTF-8 --detect --text .*/a#1 50%\.ppt$')
endfunction

" docx2txt.pl was fed the buffer instead of the file.
function! Test_docx2txt_pl_reads_file() abort
  call s:Tools('docx2txt.pl')
  call s:Start()
  call s:Match(s:Open(s:File('a.docx', 'raw')), '^STUB docx2txt.pl .*a\.docx -$')
  call s:ReadOnlyView('text')
endfunction

" 'in2csv --write-sheets -e utf-8' is a usage error.
function! Test_in2csv_xlsx() abort
  call s:Tools('real:in2csv')
  call s:Start()
  call s:Match(s:Open(s:File('a.xlsx', 'hello.xlsx')), '^alpha,beta\ngamma,delta$')
endfunction

function! Test_in2csv_xls() abort
  call s:Tools('real:in2csv')
  call s:Start()
  call s:Match(s:Open(s:File('a.xls', 'hello.xls')), '^alpha,beta$')
endfunction

" *.xltx and *.xltm were left to the zip plug-in, so the converter never ran.
function! Test_xltx_is_converted_not_zip_browsed() abort
  call s:Tools('xlsx2csv')
  call s:Start('zip')
  call s:Match(s:Open(s:File('a.xltx', 'hello.xlsx')), '^STUB xlsx2csv')
endfunction

" ':filetype off' removed the autocommands and ':filetype on' did not restore them.
function! Test_filetype_off_on() abort
  call s:Tools('pdftotext')
  call s:Start()
  filetype off
  filetype on
  call s:Match(s:Open(s:File('a.pdf', 'raw')), '^STUB pdftotext')
endfunction

" xlsx2csv changed the global 'fileformat'.
function! Test_xlsx_keeps_global_fileformat() abort
  call s:Tools('xlsx2csv')
  call s:Start()
  setglobal fileformat=dos
  call s:Open(s:File('a.xlsx', 'raw'))
  call s:Check(&g:fileformat ==# 'dos', 'global fileformat is ' . &g:fileformat)
  call s:Check(&l:fileformat ==# 'unix', 'local fileformat is ' . &l:fileformat)
endfunction

" Converter output without filetype, such as Markdown starting with #, got
" the filetype 'conf' from the fallback detection, which undid the read-only view.
function! Test_markitdown_xlsx_stays_readonly() abort
  call s:Tools('markitdown')
  call s:Start()
  call s:Match(s:Open(s:File('a.xlsx', 'raw')), '^## STUB markitdown')
  call s:ReadOnlyView('markdown')
endfunction

function! Test_markitdown_pdf_is_markdown() abort
  call s:Tools('markitdown')
  call s:Start()
  call s:Match(s:Open(s:File('a.pdf', 'raw')), '^## STUB markitdown')
  call s:ReadOnlyView('markdown')
endfunction

" Any other output without filetype becomes 'text'.
function! Test_xlsx_tika_html_is_text() abort
  call s:Tools('tika', 'lynx')
  call s:Start()
  call s:Open(s:File('a.xlsx', 'raw'))
  call s:ReadOnlyView('text')
endfunction

" Unchanged behavior {{{1

function! Test_disabled_by_loaded_office() abort
  call s:Tools('pdftotext')
  let g:loaded_office = 1
  call s:Start()
  call s:Match(s:Open(s:File('a.pdf', 'raw')), '^raw$')
endfunction

function! Test_no_duplicate_autocommands_when_sourced_twice() abort
  call s:Tools('pdftotext')
  call s:Start()
  runtime! ftdetect/office.vim
  redir => out
  silent autocmd BufReadPost *.pdf
  redir END
  call s:Check(len(split(out, 'pdf()', 1)) == 2, 'autocommands: ' . out)
endfunction

function! Test_docx_zip_fallback_without_converter() abort
  call s:Tools()
  call s:Start('zip')
  call s:Match(s:Open(s:File('a.docx', 'hello.docx')), 'word/document\.xml')
endfunction

function! Test_pdf_real() abort
  call s:Tools('real:pdftotext')
  call s:Start()
  call s:Match(s:Open(s:File('a.pdf', 'hello.pdf')), 'Hello from odt')
  call s:ReadOnlyView('text')
endfunction

function! Test_docx_real_docx2txt() abort
  call s:Tools('real:docx2txt')
  call s:Start()
  call s:Match(s:Open(s:File('a.docx', 'hello.docx')), 'Hello from odt')
  call s:ReadOnlyView('text')
endfunction

function! Test_docx_real_pandoc() abort
  call s:Tools('real:pandoc')
  call s:Start()
  call s:Match(s:Open(s:File('a.docx', 'hello.docx')), 'Hello from odt')
  call s:ReadOnlyView('markdown')
endfunction

function! Test_odt_real_pandoc() abort
  call s:Tools('real:pandoc')
  call s:Start()
  call s:Match(s:Open(s:File('a.odt', 'hello.odt')), 'Hello from odt')
  call s:ReadOnlyView('text')
endfunction

function! Test_xlsx_real_xlsx2csv() abort
  call s:Tools('real:xlsx2csv')
  call s:Start()
  call s:Match(s:Open(s:File('a.xlsx', 'hello.xlsx')), "alpha\tbeta")
  call s:ReadOnlyView('tsv')
  call s:Check(!&l:wrap && &l:list, 'tab stops not shown')
endfunction

function! Test_ods_real_soffice() abort
  call s:Tools('real:soffice')
  call s:Start()
  call s:Match(s:Open(s:File('a b.ods', 'hello.ods')), '^alpha,beta\ngamma,delta$')
  call s:ReadOnlyView('csv')
endfunction

" }}}

function! s:Run(name) abort
  let result = 'PASS'
  try
    call call(a:name, [])
  catch /^SKIP: /
    let result = v:exception
  catch
    call add(s:errors, v:exception . ' @ ' . v:throwpoint)
  endtry
  if !empty(s:errors)
    let result = 'FAIL: ' . join(s:errors, '; ')
  endif
  call writefile([result], $OFFICE_TEST_RESULT)
  call delete(s:tmp, 'rf')
endfunction

call s:Run(g:test)
qall!

" vim: set foldmethod=marker:
