if exists('g:loaded_office') || &cp
  finish
endif
let g:loaded_office = 1

let s:keepcpo         = &cpo
set cpo&vim

let s:slash  = exists('+shellslash') && !&shellslash ? '\' : '/'
let s:tmpdir = has('win32') ? $TMP : $TMPDIR
let s:cat    = has('win32') ? 'type' : 'cat'
let s:nul    = has('win32') ? 'NUL' : '/dev/null'

if executable('w3m')
  let  s:browser = 'w3m -dump -T text/html'
elseif executable('elinks')
  let  s:browser = 'elinks -dump 1'
elseif executable('lynx')
  let  s:browser = 'lynx -dump -stdin'
endif

" Remove extensions jar?|epub|doc[xm]|xls[xmb]|pp[st][xm] from g:zipPlugin_ext
" from Sep 13, 2016 and afterwards add back whenever converter unavailable
" NOTE: adding back might happen too late for zipPlugin to be aware of it!
let g:zipPlugin_ext='*.apk,*.celzip,*.crtx,*.ear,*.gcsx,*.glox,*.gqsx,*.kmz,*.oxt,*.potm,*.potx,*.ppam,*.sldx,*.thmx,*.vdw,*.war,*.wsz,*.xap,*.xlam,*.xlam,*.xltm,*.xltx,*.xpi,*.zip'

" {{{ PDF
autocmd BufReadPost *.pdf call s:pdf()
function! s:pdf()
  if executable('pdftotext') && exists('s:browser')
    silent exe '%!pdftotext -htmlmeta -enc UTF-8 ' . expand('%:p:S') . ' -' . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('pdftotext')
    silent %!pdftotext -enc UTF-8 %:p:S -
    setlocal fileencoding=utf-8
  elseif executable('pdftohtml') && exists('s:browser')
    silent exe '%!pdftohtml -i -noframes -nodrm -enc UTF-8 -stdout ' . expand('%:p:S') . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika') && exists('s:browser')
    silent exe '%!tika --encoding=UTF-8 --detect --html ' . expand('%:p:S') . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -
    setlocal fileencoding=utf-8
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endfunction
" }}}

" {{{ EPUB
autocmd BufReadPost *.epub call s:epub()
function! s:epub()
  if executable('pandoc')
    silent %!pandoc --from=epub --to=plain %:p:S
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -
    setlocal fileencoding=utf-8
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
  let exts = '*.epub'
  let g:zipPlugin_ext .= ',' .  exts
  if exists('#zip') && !exists('#zip#BufReadCmd#' . exts)
    exe 'au zip BufReadCmd ' . exts . ' call zip#Browse(expand("<amatch>"))'
    doautocmd zip BufReadCmd
  endif
endfunction
" }}}

" {{{ RTF
autocmd BufReadPost *.rtf call s:rtf()
function! s:rtf()
  if executable('unrtf')
    silent %!unrtf --text
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -
    setlocal fileencoding=utf-8
  elseif executable('soffice')
    silent %!soffice --cat %:p:S
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endfunction
" }}}

" {{{ ODT
autocmd BufReadPost *.odt call s:odt()
function! s:odt()
  if executable('odt2txt')
    silent %!odt2txt --encoding=UTF-8 %:p:S
    setlocal fileencoding=utf-8
  elseif executable('odt2html') && exists('s:browser')
    silent exe '%!odt2txt --encoding=UTF-8 ' . expand('%:p:S') . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('pandoc')
    silent %!pandoc --from=odt --to=plain %:p:S
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -
    setlocal fileencoding=utf-8
  elseif executable('soffice')
    silent %!soffice --cat %:p:S
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endfunction
" }}}

" {{{ DOC
autocmd BufReadPost *.do{c,t} call s:doc()
function! s:doc()
  if executable('wvText')
    silent %!wvText %:p:S
  elseif executable('wvHtml') && exists('s:browser')
    silent exe '%!wvText ' . expand('%:p:S') . ' | ' . s:browser
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
  elseif executable('soffice')
    silent %!soffice --cat %:p:S
  elseif executable('catdoc')
    silent %!catdoc %:p:S
  elseif executable('antiword')
    silent %!antiword %:p:S
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endfunction
" }}}

" {{{ DOCX
autocmd BufReadPost *.do{c,t}{m,x} call s:docx()
function! s:docx()
  if executable('pandoc')
    silent %!pandoc --from=docx --to=plain %:p:S
    setlocal fileencoding=utf-8
  elseif executable('docx2txt.pl')
    silent %!docx2txt.pl -'
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -
    setlocal fileencoding=utf-8
  elseif executable('soffice')
    silent %!soffice --cat %:p:S
  else
    let exts = '*.docx,*.docm,*.dotx,*.dotm'
    let g:zipPlugin_ext .= ',' .  exts
    if exists('#zip') && !exists('#zip#BufReadCmd#' . exts)
      exe 'au zip BufReadCmd ' . exts . ' call zip#Browse(expand("<amatch>"))'
      doautocmd zip BufReadCmd
    endif
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endfunction
" }}}

" {{{ XLS
autocmd BufReadPost *.xl{m,s,t} call s:xls()
function! s:xls()
  if executable('xlhtml') && exists('s:browser')
    silent silent exe '%!xlhtml ' . expand('%:p:S') . ' | ' . s:browser
  elseif executable('excel2csv')
    silent exe '%!excel2csv --trim ' . expand('%:p:S')
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */,/' | endif
    setlocal filetype=csv
  elseif executable('xls2csv')
    let in = expand('%:p:S')
    let out = s:tmpdir . s:slash . expand('%:t:r')
    silent exe '%!xls2csv -q -a UTF-8 -b WINDOWS-1252 -x ' . in . ' -c ' . out . ' > ' . s:nul . ' && ' . s:cat . ' ' . out
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */,/' | endif
    setlocal filetype=csv
  elseif executable('soffice') && exists('s:browser')
    silent exe '%!soffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . expand('%:p:S') . ' > ' . s:nul . ' && ' . s:cat . ' ' . s:tmpdir . s:slash . expand('%:t:r') . '.html' 
  elseif executable('soffice')
    silent exe '%!soffice --headless --convert-to csv --outdir ' s:tmpdir . ' ' . expand('%:p:S') . ' > ' . s:nul . ' && ' . s:cat . ' ' . s:tmpdir . s:slash . expand('%:t:r') . '.csv' 
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */,/' | endif
    setlocal filetype=csv
  elseif executable('tika') && exists('s:browser')
    silent exe '%!tika --encoding=UTF-8 --detect --html ' . expand('%:p:S') . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -
    setlocal fileencoding=utf-8
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */\t\+/' | endif
  endif
  setlocal nowrap
  setlocal nomodifiable readonly
endfunction
" }}}

" {{{ XLSX
autocmd BufReadPost *.xl{s,t}{x,m,b} call s:xlsx()
function! s:xlsx()
  if executable('git-xlsx-textconv.pl')
    silent %!git-xlsx-textconv.pl %:p:S
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */\t\+/' | endif 
    setlocal filetype=text
  elseif executable('git-xlsx-textconv')
    silent %!git-xlsx-textconv %:p:S
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */\t\+/' | endif
    setlocal filetype=text
  elseif executable('excel2csv')
    silent %!excel2csv %:p:S
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */,/' | endif 
    setlocal filetype=csv
  elseif executable('soffice') && exists('s:browser')
    silent exe '%!soffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . expand('%:p:S') . ' > ' . s:nul . ' && ' . s:cat . ' ' . s:tmpdir . s:slash . expand('%:t:r') . '.html' . ' | ' . s:browser
  elseif executable('soffice')
    silent exe '%!soffice --headless --convert-to csv --outdir ' s:tmpdir . ' ' . expand('%:p:S') . ' > ' . s:nul . ' && ' . s:cat . ' ' . s:tmpdir . s:slash . expand('%:t:r') . '.csv'
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */,/' | endif
    setlocal filetype=csv
  elseif executable('tika') && exists('s:browser')
    silent exe '%!tika --encoding=UTF-8 --detect --html ' . expand('%:p:S') . ' | ' . s:browser
    setlocal fileencoding=utf-8
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */\t\+/' | endif
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
    if exists(':EasyAlign') == 2 | exe '%EasyAlign */\t\+/' | endif
    setlocal filetype=text
  else
    let exts = '*.xlsx,*.xlsm,*.xlsb'
    let g:zipPlugin_ext .= ',' .  exts
    if exists('#zip') && !exists('#zip#BufReadCmd#' . exts)
      exe 'au zip BufReadCmd ' . exts . ' call zip#Browse(expand("<amatch>"))'
      doautocmd zip BufReadCmd
    endif
  endif
  setlocal nowrap
  setlocal nomodifiable readonly
endfunction
" }}}

" {{{ PPT
autocmd BufReadPost *.pp{s,t} call s:ppt()
function! s:ppt()
  if executable('tika') && exists('s:browser') 
    silent exe '%!tika --encoding=UTF-8 --detect --html -' . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent %!tika --encoding=UTF-8 --detect --text -
    setlocal fileencoding=utf-8
  elseif executable('soffice') && exists('s:browser')
    silent exe '%!soffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . expand('%:p:S') . ' > ' . s:nul . ' && ' . s:cat . ' ' . s:tmpdir . s:slash . expand('%:t:r') . '.html' . ' | ' . s:browser
  elseif executable('soffice')
    silent %!soffice --cat %:p:S
  elseif executable('catppt')
    silent %!catppt %:p:S
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endfunction
" }}}

" {{{ PPTX
autocmd BufReadPost *.pp{s,t}{x,m} call s:pptx()
function! s:pptx()
  if executable('pptx2md')
    let output_file = s:tmpdir . s:slash . expand('%:t:r') . '.md'
    silent exe '%!pptx2md --disable_image --disable_wmf ' . expand('%:p:S') . ' -o ' . output_file . ' > ' . s:nul . ' && ' . s:cat . ' ' output_file
    setlocal filetype=markdown foldmethod=expr foldexpr=MarkdownFold()
  elseif executable('tika') && exists('s:browser') 
    silent exe '%!tika --encoding=UTF-8 --detect --html -' . ' | ' . s:browser
    setlocal fileencoding=utf-8
    setlocal filetype=text
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
    setlocal filetype=text
  elseif executable('soffice') && exists('s:browser')
    silent exe '%!soffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . expand('%:p:S') . ' > ' . s:nul . ' && ' . s:cat . ' ' . s:tmpdir . s:slash . expand('%:t:r') . '.html' . ' | ' . s:browser
  elseif executable('soffice')
    silent %!soffice --cat %:p:S
    setlocal filetype=text
  else
    let exts = '*.pptx,*.pptm,*.ppsx,*.ppsm'
    let g:zipPlugin_ext .= ',' .  exts
    if exists('#zip') && !exists('#zip#BufReadCmd#' . exts)
      exe 'au zip BufReadCmd ' . exts . ' call zip#Browse(expand("<amatch>"))'
      doautocmd zip BufReadCmd
    endif
  endif
  setlocal nomodifiable readonly
endfunction
" }}}

" {{{ OTHER FILE TYPES
if executable('tika')
  autocmd BufReadPost *.{{rar,7z},{class,ja,jar},chm,{mdb,accdb},{pst,msg},mht{,ml}}
        \ silent exe '%!tika --encoding=UTF-8 --detect --text -' |
        \ setlocal fileencoding=utf-8 |
        \ setlocal filetype=text nomodifiable readonly 
else
  let exts = '*.ja,*.jar'
  let g:zipPlugin_ext .= ',' .  exts
  if exists('#zip') && !exists('#zip#BufReadCmd#' . exts)
    exe 'au zip BufReadCmd ' . exts . ' call zip#Browse(expand("<amatch>"))'
    doautocmd zip BufReadCmd
  endif
endif
" }}}

" ------------------------------------------------------------------------------
let &cpo= s:keepcpo
unlet s:keepcpo

" ex: set foldmethod=marker: 
