This Vim plug-in lets Vim read the text of common binary files such as those of type `PDF`, `Word` (`doc(x)`), `Excel` (`xls(x)`), `Powerpoint` (`ppt(x)`), Open Document (`odt`), `zip`...

It depends on appropriate external converters such as `pdftotext`, `odt2txt`, `docx2txt.pl`, `pandoc`, `pptx2md`, ... to do so.
For an all-in-one solution:

1. Download the latest runnable `tika-app-...jar` from the [Apache project](https://tika.apache.org/download.html) to `~/bin/tika.jar`.

0. Create a shell script `~/bin/tika` that reads

    ```sh
        #!/bin/sh
        exec java -jar "$HOME/bin/tika.jar" "$@" 2>/dev/null
    ```
    and mark it executable (by `chmod a+x ~/bin/tika`).

0. Add `~/bin` to your `$PATH` variable; for example by adding to `~/.profile` or `~/.zshenv` the line

    ```sh
        PATH=$PATH:~/bin
    ```

