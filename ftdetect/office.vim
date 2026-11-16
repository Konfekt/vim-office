" Read PDF, office documents and other binary files as text.
" The conversions are in autoload/office.vim.
" Like all ftdetect scripts, this one is sourced inside the filetypedetect augroup.
" vint: -ProhibitAutocmdWithNoGroup

" Skip if disabled by the user, or if the autocommands are still defined;
" ':filetype off' removes them, so ':filetype on' must define them again.
if &cp || exists('g:loaded_office') && !exists('s:loaded') || exists('#User#OfficeLoaded')
  finish
endif
let g:loaded_office = 1
let s:loaded = 1

" Marks that the autocommands below are defined.
autocmd User OfficeLoaded :

" Remove the extensions jar?|epub|doc[xm]|xls[xmb]|xlt[xm]|pp[st][xm] from g:zipPlugin_ext,
" so that the converters below handle them; they are added back whenever a converter is unavailable.
" NOTE: adding back might happen too late for zipPlugin to be aware of it!
let g:zipPlugin_ext = '*.apk,*.celzip,*.crtx,*.ear,*.gcsx,*.glox,*.gqsx,*.kmz,*.oxt,*.potm,*.potx,*.ppam,*.sldx,*.thmx,*.vdw,*.war,*.wsz,*.xap,*.xlam,*.xpi,*.zip'

autocmd BufReadPost *.pdf               call office#pdf()
autocmd BufReadPost *.djvu,*.djv        call office#djvu()
autocmd BufReadPost *.epub              call office#epub()
autocmd BufReadPost *.rtf               call office#rtf()
autocmd BufReadPost *.{odt,sxw}         call office#odt()
autocmd BufReadPost *.ods               call office#ods()
autocmd BufReadPost *.odp               call office#odp()
autocmd BufReadPost *.doc               call office#doc()
autocmd BufReadPost *.do{c,t}{m,x}      call office#docx()
autocmd BufReadPost *.xl{m,s,t}         call office#xls()
autocmd BufReadPost *.xl{s,t}{x,m,b}    call office#xlsx()
autocmd BufReadPost *.pp{s,t}           call office#ppt()
autocmd BufReadPost *.pp{s,t}{x,m}      call office#pptx()

if executable('tika')
  autocmd BufReadPost *.{{rar,7z},{class,ja,jar},chm,{mdb,accdb},{pst,msg},mht{,ml}} call office#tika()
else
  let g:zipPlugin_ext .= ',*.ja,*.jar'
  for s:pat in ['*.ja', '*.jar']
    if exists('#zip') && !exists('#zip#BufReadCmd#' . s:pat)
      execute 'autocmd zip BufReadCmd ' . s:pat . ' call zip#Browse(expand("<amatch>"))'
    endif
  endfor
  unlet s:pat
endif

if executable('tesseract')
  " TODO: convert 2-character v:lang to 3-character ISO 639-2 language code
  " by list from https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes
  " and check whether it is in systemlist('tesseract --list-langs').
  " Then pass to tesseract by appending '-l eng' . '+' . lang
  autocmd BufReadPost *.{{jpg,jpeg},png,gif,{tif,tiff},webp,heif,raw,bmp,psd,indd} call office#tesseract()
endif
