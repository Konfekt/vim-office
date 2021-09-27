This plug-in enables Vim to read text documents of type `PDF`, `Microsoft Office` such as `Word` (`doc(x)`), `Excel` (`xls(x)`) or `Powerpoint` (`ppt(x)`), Open Document (`odt`), `EPUB` ....
[To go even further, for example, to read, among many others file formats, media files in Vim, add this [Vimscript  snippet](https://github.com/wofr06/lesspipe/wiki/vim) from the Wiki for [lesspipe.sh](https://github.com/wofr06/lesspipe/) to your `vimrc`!]

For the text extraction, it depends on external tools, but most use cases are covered by an installation of

  - `LibreOffice` and a common text browser (such as `lynx`), and
  - `pdftotext`.

It uses, whenever available, appropriate external converters such as [unrtf](http://ftp.gnu.org/gnu/unrtf/), [pandoc](http://pandoc.org), [docx2txt.pl](https://github.com/arthursucks/docx2txt), [odt2txt](https://github.com/dstosberg/odt2txt), [xlscat](https://github.com/Tux/Spreadsheet-Read/tree/master/scripts), [xlsx2csv.py](https://github.com/dilshod/xlsx2csv) or [pptx2md](https://github.com/ssine/pptx2md) ..., but will fall back to:

- Either [LibreOffice](https://www.libreoffice.org/download/download/) which is an office suite that (together with a common text browser such as `lynx`) can handle all those formats listed above, except `PDF`s.
    (To use it on Microsoft Windows, ensure after its installation that its path is added to the `%PATH%` environment variable, say by [Rapidee](http://www.rapidee.com/).)
- Or [Tika](https://tika.apache.org/download.html) which is a content extractor that can handle all those formats listed above and many more.
    To use it:

    1. Download the latest runnable `tika-app-...jar` from [Tika](https://tika.apache.org/download.html) to `~/bin/tika.jar` (on Linux) respectively `%USERPROFILE%\bin` (on Microsoft Windows).

    0. Create

        - on Linux, a shell script `~/bin/tika` that reads
        ```sh
            #!/bin/sh
            exec java -Dfile.encoding=UTF-8 -jar "$HOME/bin/tika.jar" "$@" 2>/dev/null
        ```
        and mark it executable (by `chmod a+x ~/bin/tika`).

        - on Microsoft Windows, a batch script `%USERPROFILE%\bin\tika.bat` that reads
        ```bat
            @echo off
            java -Dfile.encoding=UTF-8 -jar "%USERPROFILE%\bin\tika.jar" %*
        ```

    0. Add the folder of the newly created `tika` executable to your environment variable `$PATH` (on Linux) respectively `%PATH%` (on Microsoft Windows):

        - on Linux, if you use `bash` or `zsh` by adding to `~/.profile` or `~/.zshenv` the line

        ```sh
            PATH=$PATH:~/bin
        ```

        - on Microsoft Windows, a convenient program to update `%PATH%` is [Rapidee](http://www.rapidee.com/).

