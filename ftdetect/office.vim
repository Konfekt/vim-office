autocmd BufReadPost *.{pdf,odt,rtf,do{c,t}{,m,x},xls{,x,m},ppt{,x,m}} setlocal filetype=text

" {{{ PDF
if executable('pdftotext')
  autocmd BufReadPost *.pdf silent exe '%!pdftotext ' . shellescape(expand('%')) . ' -'
elseif executable('tika')
  autocmd BufReadPost *.pdf silent exe '%!tika --text ' . shellescape(expand('%'))
endif
" }}}

" {{{ ODT
if executable('odt2txt')
  autocmd BufReadPost *.odt silent exe '%!odt2txt ' . shellescape(expand('%'))
elseif executable('pandoc')
  autocmd BufReadPost *.odt silent exe '%!pandoc --from=odt --to=plain ' . shellescape(expand('%'))
elseif executable('tika')
  autocmd BufReadPost *.odt silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.odt  call s:unoconv('txt')
elseif executable('libreoffice')
  autocmd BufReadPost *.odt silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
endif
" }}}

" {{{ RTF
if executable('unrtf')
  autocmd BufReadPost *.rtf silent %!unrtf -P /etc/unrtf --text
elseif executable('tika')
  autocmd BufReadPost *.rtf silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.rtf call s:unoconv('txt')
elseif executable('libreoffice')
  autocmd BufReadPost *.rtf silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
endif
" }}}

" {{{ DOC
if executable('wvText')
  autocmd BufReadPost *.do{c,t} silent exe '%!wvText ' . shellescape(expand('%'))
elseif executable('tika')
  autocmd BufReadPost *.do{c,t} silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.do{c,t}  call s:unoconv('txt')
elseif executable('libreoffice')
  autocmd BufReadPost *.do{c,t} silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
elseif executable('catdoc')
  autocmd BufReadPost *.do{c,t} silent exe '%!catdoc ' . shellescape(expand('%'))
endif
" }}}

" {{{ DOCX
if executable('pandoc')
  autocmd BufReadPost *.do{c,t}{m,x} silent exe '%!pandoc --from=docx --to=plain ' . shellescape(expand('%'))
elseif executable('docx2txt.pl')
  autocmd BufReadPost *.do{c,t}{m,x} silent exe '%!docx2txt.pl -'
elseif executable('tika')
  autocmd BufReadPost *.do{c,t}{m,x} silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.do{c,t}{m,x} call s:unoconv('txt')
elseif executable('libreoffice')
  autocmd BufReadPost *.do{c,t}{m,x} silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
endif
" }}}

" {{{ XLS
if executable('tika')
  autocmd BufReadPost *.xls
        \ silent exe '%!tika --text ' . shellescape(expand('%')) |
        \ if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif |
        \ setlocal nowrap
elseif executable('xls2csv')
  autocmd BufReadPost *.xls
        \ silent silent exe '%!xls2csv ' . shellescape(expand('%'))|
        \ if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif |
        \ setlocal nowrap nowrap filetype=csv
elseif executable('unoconv')
  autocmd BufReadPost *.xls
        \ call s:unoconv('csv') |
        \ if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif |
        \ setlocal nowrap filetype=csv
elseif executable('libreoffice')
  autocmd BufReadPost *.xls
        \ silent exe '%!libreoffice --convert-to csv ' . shellescape(expand('%'))
        \ if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif |
        \ setlocal nowrap nowrap filetype=csv
endif
" }}}

" {{{ XLSX
if executable('git-xlsx-textconv.pl')
  autocmd BufReadPost *.xls{x,m,b}
        \ silent silent exe '%!git-xlsx-textconv.pl ' . shellescape(expand('%'))|
        \ if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif |
        \ setlocal nowrap
elseif executable('git-xlsx-textconv')
  autocmd BufReadPost *.xls{x,m,b} 
        \ silent silent exe '%!git-xlsx-textconv ' . shellescape(expand('%'))|
        \ if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif |
        \ setlocal nowrap
elseif executable('tika')
  autocmd BufReadPost *.xls{x,m,b}
        \ silent exe '%!tika --text ' . shellescape(expand('%')) |
        \ if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif |
        \ setlocal nowrap
elseif executable('unoconv')
  autocmd BufReadPost *.xls{x,m,b}
        \ call s:unoconv('csv') |
        \ if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif |
        \ setlocal nowrap setlocal filetype=csv
elseif executable('libreoffice')
  autocmd BufReadPost *.xls{x,m,b}
        \ silent exe '%!libreoffice --convert-to csv ' . shellescape(expand('%'))
        \ if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif |
      \ setlocal nowrap nowrap filetype=csv
endif
" }}}

" {{{ PPT
if executable('tika')
  autocmd BufReadPost *.ppt silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.ppt call s:unoconv('txt')
elseif executable('libreoffice')
  autocmd BufReadPost *.ppt silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
elseif executable('catppt')
  autocmd BufReadPost *.ppt silent exe '%!catppt ' . shellescape(expand('%'))
endif
" }}}

" {{{ PPTX
if executable('pptx2md')
  autocmd BufReadPost *.ppt{x,m}
        \ silent exe '%!pptx2md --disable_image --disable_wmf ' . shellescape(expand('%')) . ' -o ' . $TMPDIR . '/presentation.md >/dev/null 2>&1 && cat ' . $TMPDIR . '/presentation.md'|
        \ setlocal filetype=markdown
elseif executable('tika')
  autocmd BufReadPost *.ppt{x,m} silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.ppt{x,m} call s:unoconv('txt')
elseif executable('libreoffice')
  autocmd BufReadPost *.ppt{x,m} silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
endif
" }}}

" {{{ OTHER
if executable('tika')
  autocmd BufReadPost *.{rar,7z,{mdb,accdb},{pst,msg},{class,jar},epub,chm,mht{,ml},pps{,x,m}}
        \ silent exe '%!tika --text ' . shellescape(expand('%')) |
        \ setlocal filetype=text nomodifiable readonly
endif
" }}}

autocmd BufReadPost *.{pdf,odt,rtf,do{c,t}{,m,x},xls{,x,m,b},ppt{,x,m}} setlocal nomodifiable readonly

function! s:unoconv(ft)
  let tmp=tempname() . '.' . a:ft
  silent exe '!unoconv --output=' . tmp . ' ' . shellescape(expand('%'))
  bwipe! | exe 'edit! ' . tmp
endfunction

" ex: set foldmethod=marker: 
