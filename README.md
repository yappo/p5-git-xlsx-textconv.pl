# git-xlsx-textconv.pl - git text converter for xlsx file

git diff wrapper for xlsx file.

<img src="http://t.co/5Epi6NXHZ5">

# install

    cpanm Spreadsheet::XLSX
    cp git-xlsx-textconv.pl $PATH/
    
# .gitconfig

    [diff "xlsx"]
        binary = true
        textconv = $PATH/git-xlsx-textconv.pl

# .gitattributes in your project repository

    *.xlsx diff=xlsx

# LICENSE

Copyright (C) Kazuhiro Osawa 2014-

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

