This plug-in lets Vim read documents of type `PDF`, `Word` (`doc(x)`), `Excel` (`xls(x)`), `Powerpoint` (`ppt(x)`), Open Document (`odt`), ... via `LibreOffice`, `Tika` or specific converters.

It uses, whenever available, appropriate external converters such as `pdftotext`, `odt2txt`, `docx2txt.pl`, `pandoc`, `pptx2md`, ... to do so, but falls back to:

- Either [LibreOffice](https://www.libreoffice.org/download/download/) which is an office suite that can handle all those formats listed above, except `PDF`s.
    (To use it on Microsoft Windows, ensure after its installation that its path is added to the `%PATH%` environment variable, say by [Rapidee](http://www.rapidee.com/).)
- Or [Tika](https://tika.apache.org/download.html) which is a content extractor that can handle all those formats listed above and many more.
    To use it:

    1. Download the latest runnable `tika-app-...jar` from [Tika](https://tika.apache.org/download.html) to `~/bin/tika.jar` (on Linux) respectively `%USERPROFILE%\bin` (on Microsoft Windows).

    0. Create

        - on Linux, a shell script `~/bin/tika` that reads
        ```sh
            #!/bin/sh
            exec java -jar "$HOME/bin/tika.jar" "$@" 2>/dev/null
        ```
        and mark it executable (by `chmod a+x ~/bin/tika`).

        - on Microsoft Windows, a batch script `%USERPROFILE%\bin\tika.bat` that reads
        ```bat
            @echo off
            java -jar "%USERPROFILE%\bin\tika.jar" %*
        ```

    0. Add the folder of the newly created `tika` executable to your environment variable `$PATH` (on Linux) respectively `%PATH%` (on Microsoft Windows):

        - on Linux, if you use `bash` or `zsh` by adding to `~/.profile` or `~/.zshenv` the line

        ```sh
            PATH=$PATH:~/bin
        ```

        - on Microsoft Windows, a convenient program to update `%PATH%` is [Rapidee](http://www.rapidee.com/).

