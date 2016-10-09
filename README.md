*Vim-Office*
========

This plug-in makes Vim read the text of common binary files such as those of type `PDF`, Word (`doc(x)`), Powerpoint (`ppt(x)`), Open Document (`odt`), `zip`...

It depends on programs such as `pdftotext`, `odt2txt`, ... to do so. For an
all-in-one solution,

- add a executable shell script `~/bin/tika` with content

```sh
    #!/bin/sh
    exec java -jar "$HOME/bin/tika.jar" "$@" 2>/dev/null
```

- add `~/bin` to your `$PATH` variable; for example by adding to `~/.profile` or `~/.zshenv` the line
```sh
    PATH=~/bin:$PATH
```

- download `tika.jar` from the Apache project to `~/bin/tika.jar`.

