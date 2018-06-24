*Vim-Office*
========

This plug-in makes Vim read the text of common binary files such as those of type `PDF`, `Word` (`doc(x)`), `Excel` (`xls(x)`), `Powerpoint` (`ppt(x)`), Open Document (`odt`), `zip`...

It depends on programs such as `pdftotext`, `odt2txt`, ... to do so.
For an all-in-one solution:

- Download the latest runnable `tika-app-...jar` from the [Apache project](https://tika.apache.org/download.html) to `~/bin/tika.jar`.

- Create a shell script `~/bin/tika` that reads

    ```sh
        #!/bin/sh
        exec java -jar "$HOME/bin/tika.jar" "$@" 2>/dev/null
    ```
    and mark it executable (by `chmod a+x ~/bin/tika`).

- Add `~/bin` to your `$PATH` variable; for example by adding to `~/.profile` or `~/.zshenv` the line
```sh
    PATH=$PATH:~/bin
```


