if executable('pdftotext')
  autocmd BufReadPost *.pdf silent exe '%!pdftotext ' . shellescape(expand('%')) . ' -'
elseif executable('tika')
  autocmd BufReadPost *.pdf silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.pdf  call s:unoconv('txt')
endif

if executable('odt2txt')
  autocmd BufReadPost *.odt silent exe '%!odt2txt ' . shellescape(expand('%'))
elseif executable('tika')
  autocmd BufReadPost *.odt silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.odt  call s:unoconv('txt')
endif

if executable('docx2txt.pl')
  autocmd BufReadPost *.do{c,t}{m,x} silent exe '%!docx2txt.pl -'
elseif executable('tika')
  autocmd BufReadPost *.do{c,t}{m,x} silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.do{c,t}{m,x} call s:unoconv('txt')
endif

if executable('wvText')
  autocmd BufReadPost *.do{c,t} silent exe '%!wvText ' . shellescape(expand('%'))
elseif executable('tika')
  autocmd BufReadPost *.do{c,t} silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.do{c,t}  call s:unoconv('txt')
endif

if executable('unrtf')
  autocmd BufReadPost *.rtf silent %!unrtf -P /etc/unrtf --text
elseif executable('tika')
  autocmd BufReadPost *.rtf silent exe '%!tika --text ' . shellescape(expand('%'))
elseif executable('unoconv')
  autocmd BufReadPost *.rtf call s:unoconv('txt')
endif

autocmd BufReadPost *.{pdf,odt,docx,rtf} setlocal filetype=text nomodifiable readonly

if executable('tika')
  autocmd BufReadPost *.{rar,7z,{mdb,accdb},{pst,msg},{class,jar},epub,chm,mht{,ml},ppt{,x,m},pps{,x,m}}
        \ silent exe '%!tika --text ' . shellescape(expand('%')) |
        \ setlocal filetype=text nomodifiable readonly
  autocmd BufReadPost *.xls{,x,m,b}
        \ silent exe '%!tika --text ' . shellescape(expand('%')) |
        \ if exists(':EasyAlign') | exe '%EasyAlign */\t\+/' | endif |
        \ setlocal filetype=text nomodifiable readonly nowrap
elseif executable('unoconv')
  autocmd BufReadPost *.{xls{,x,m,b}}
        \ call s:unoconv('csv') |
        \ setlocal filetype=csv nomodifiable readonly
endif

function! s:unoconv(ft)
  let tmp=tempname() . '.' . a:ft
  silent exe '!unoconv --output=' . tmp . ' ' . shellescape(expand('%'))
  bwipe! | exe 'edit! ' . tmp
endfunction
