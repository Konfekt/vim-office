let s:tmpdir = has('win32') ? $TMP : $TMPDIR

if executable('w3m')
  let  s:browser = 'w3m -dump -T text/html'
elseif executable('elinks')
  let  s:browser = 'elinks -dump 1'
elseif executable('lynx')
  let  s:browser = 'lynx -dump -stdin'
endif

" {{{ PDF
autocmd BufReadPost *.pdf call s:pdf()
fun s:pdf()
  if executable('pdftotext') && exists('s:browser')
    silent exe '%!pdftotext -htmlmeta -enc UTF-8 ' . shellescape(expand('%')) . ' -' . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('pdftotext')
    silent exe '%!pdftotext -enc UTF-8 ' . shellescape(expand('%')) . ' -'
    setlocal fileencoding=utf-8
  elseif executable('pdftohtml') && exists('s:browser')
    silent exe '%!pdftohtml -i -noframes -nodrm -enc UTF-8 -stdout ' . shellescape(expand('%')) . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika') && exists('s:browser')
    silent exe '%!tika --encoding=UTF-8 --detect --html' . shellescape(expand('%')) . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endf
" }}}

" {{{ ODT
autocmd BufReadPost *.odt call s:odt()
fun s:odt()
  if executable('odt2txt')
    silent exe '%!odt2txt --encoding=UTF-8 ' . shellescape(expand('%'))
    setlocal fileencoding=utf-8
  elseif executable('odt2html') && exists('s:browser')
    silent exe '%!odt2txt --encoding=UTF-8 ' . shellescape(expand('%')) . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('pandoc')
    silent exe '%!pandoc --from=odt --to=plain ' . shellescape(expand('%'))
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
  elseif executable('libreoffice')
    silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endf
" }}}

" {{{ RTF
autocmd BufReadPost *.rtf call s:rtf()
fun s:rtf()
  if executable('unrtf')
    silent %!unrtf -P /etc/unrtf --text
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
  elseif executable('libreoffice')
    silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endf
" }}}

" {{{ DOC
autocmd BufReadPost *.do{c,t} call s:doc()
fun s:doc()
  if executable('wvText')
    silent exe '%!wvText ' . shellescape(expand('%'))
  elseif executable('wvHtml') && exists('s:browser')
    silent exe '%!wvText ' . shellescape(expand('%')) | . ' | ' . s:browser
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
  elseif executable('libreoffice')
    silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
  elseif executable('catdoc')
    silent exe '%!catdoc ' . shellescape(expand('%'))
  elseif executable('antiword')
    silent exe '%!antiword ' . shellescape(expand('%'))
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endf
" }}}

" {{{ DOCX
autocmd BufReadPost *.do{c,t}{m,x} call s:docx()
fun s:docx()
  if executable('pandoc')
    silent exe '%!pandoc --from=docx --to=plain ' . shellescape(expand('%'))
    setlocal fileencoding=utf-8
  elseif executable('docx2txt.pl')
    silent exe '%!docx2txt.pl -'
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
  elseif executable('libreoffice')
    silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endf
" }}}

" {{{ XLS
autocmd BufReadPost *.xl{m,s,t} call s:xsl()
fun s:xsl()
  if executable('xlhtml') && exists('s:browser')
    silent silent exe '%!xlhtml ' . shellescape(expand('%')) . ' | ' . s:browser
  elseif executable('xls2csv')
    silent silent exe '%!xls2csv ' . shellescape(expand('%'))|
    if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif
    setlocal filetype=csv
  elseif executable('libreoffice') && exist('s:browser')
    silent exe '%!libreoffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . shellescape(expand('%')) . ' >/dev/null && cat ' . s:tmpdir . '/' . expand('%:t:r') . '.html' 
  elseif executable('libreoffice')
    silent exe '%!libreoffice --headless --convert-to csv --outdir ' s:tmpdir . ' ' . shellescape(expand('%')) . ' >/dev/null && cat ' . s:tmpdir . '/' . expand('%:t:r') . '.csv' 
    if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif
    setlocal filetype=csv
  elseif executable('tika') && exists('s:browser')
    silent exe '%!tika --encoding=UTF-8 --detect --html ' . shellescape(expand('%')) . ' | ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
    if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif
  endif
  setlocal nowrap
  setlocal nomodifiable readonly
endf
" }}}

" {{{ XLSX
autocmd BufReadPost *.xl{s,t}{x,m,b} call s:xlsx()
fun s:xlsx()
  if executable('git-xlsx-textconv.pl')
    silent silent exe '%!git-xlsx-textconv.pl ' . shellescape(expand('%'))|
    if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif 
    setlocal filetype=text
  elseif executable('git-xlsx-textconv')
    silent silent exe '%!git-xlsx-textconv ' . shellescape(expand('%'))|
    if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif
    setlocal filetype=text
  elseif executable('tika') && exist('s:browser')
    silent exe '%!tika --encoding=UTF-8 --detect --html ' . shellescape(expand('%')) . ' | ' . s:browser
    setlocal fileencoding=utf-8
    if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif
  elseif executable('libreoffice') && exist('s:browser')
    silent exe '%!libreoffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . shellescape(expand('%')) . ' >/dev/null && cat ' . s:tmpdir . '/' . expand('%:t:r') . '.html' . ' | ' . s:browser
  elseif executable('libreoffice')
    silent exe '%!libreoffice --headless --convert-to csv --outdir ' s:tmpdir . ' ' . shellescape(expand('%')) . ' >/dev/null && cat ' . s:tmpdir . '/' . expand('%:t:r') . '.csv'
    if exists(':EasyAlign') | exe '%EasyAlign */,/' | endif
    setlocal filetype=csv
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
    if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif
    setlocal filetype=text
  endif
  setlocal nowrap
  setlocal nomodifiable readonly
endf
" }}}

" {{{ PPT
autocmd BufReadPost *.pp{s,t} call s:ppt()
fun s:ppt()
  if executable('tika') && exist('s:browser') 
    silent exe '%!tika --encoding=UTF-8 --detect --html -' . ' ' . s:browser
    setlocal fileencoding=utf-8
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
  elseif executable('libreoffice') && exist('s:browser')
    silent exe '%!libreoffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . shellescape(expand('%')) . ' >/dev/null && cat ' . s:tmpdir . '/' . expand('%:t:r') . '.html' . ' | ' . s:browser
  elseif executable('libreoffice')
    silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
  elseif executable('catppt')
    silent exe '%!catppt ' . shellescape(expand('%'))
  endif
  setlocal nomodifiable readonly
  setlocal filetype=text
endf
" }}}

" {{{ PPTX
autocmd BufReadPost *.pp{s,t}{x,m} call s:pptx()
fun s:pptx()
  if executable('pptx2md')
    let output_file = s:tmpdir . '/' . expand('%:t:r') . '.md'
    silent exe '%!pptx2md --disable_image --disable_wmf ' . shellescape(expand('%')) . ' -o ' . output_file . ' >/dev/null && cat ' output_file
    setlocal filetype=markdown foldmethod=expr foldexpr=MarkdownFold()
  elseif executable('tika') && exist('s:browser') 
    silent exe '%!tika --encoding=UTF-8 --detect --html -' . ' ' . s:browser
    setlocal fileencoding=utf-8
    setlocal filetype=text
  elseif executable('tika')
    silent exe '%!tika --encoding=UTF-8 --detect --text -'
    setlocal fileencoding=utf-8
    setlocal filetype=text
  elseif executable('libreoffice') && exist('s:browser')
    silent exe '%!libreoffice --headless --convert-to html --outdir ' s:tmpdir . ' ' . shellescape(expand('%')) . ' >/dev/null && cat ' . s:tmpdir . '/' . expand('%:t:r') . '.html' . ' | ' . s:browser
  elseif executable('libreoffice')
    silent exe '%!libreoffice --cat ' . shellescape(expand('%'))
    setlocal filetype=text
  endif
  setlocal nomodifiable readonly
endf
" }}}

" {{{ OTHER FILE TYPES
if executable('tika')
  autocmd BufReadPost *.{{rar,7z},{class,ja,jar},{epub,chm},{mdb,accdb},{pst,msg},mht{,ml}}
        \ silent exe '%!tika --encoding=UTF-8 --detect --text -' |
        \ setlocal fileencoding=utf-8 |
        \ setlocal filetype=text nomodifiable readonly 
endif
" }}}

" ex: set foldmethod=marker: 
