#!/bin/sh
# Stand-in for the external converters: tests link it under the name of the tool.
# It prints its name and arguments, or writes the file the real tool would write.
name=$(basename "$0")
case $name in
  soffice)
    fmt='' outdir='' file=''
    while [ $# -gt 0 ]; do
      case $1 in
        --convert-to) fmt=$2; shift ;;
        --outdir) outdir=$2; shift ;;
        --*) ;;
        *) file=$1 ;;
      esac
      shift
    done
    base=$(basename "$file"); base=${base%.*}
    if [ "$fmt" = html ]; then
      printf '<html><body><p>STUB soffice html</p></body></html>\n' > "$outdir/$base.$fmt"
    else
      printf 'STUB soffice %s\n' "$fmt" > "$outdir/$base.$fmt"
    fi ;;
  wvHtml)
    dir='' out=''
    for arg; do
      out=$arg
      case $arg in --targetdir=*) dir=${arg#--targetdir=} ;; esac
    done
    printf '<p>STUB wvHtml</p>\n' > "$dir/$out" ;;
  lynx)
    # Like lynx -stdin: read the HTML from standard input, ignore any file argument.
    sed 's/<[^>]*>//g' ;;
  *)
    printf 'STUB %s %s\n' "$name" "$*" ;;
esac
