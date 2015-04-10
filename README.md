# Excel Macro Extrator

Windows Command Line application to extract `.vba`-files from `.xlsm`-files.

## Usage:

~~~ sh
$ ExcelMacroExtrator.exe file targetdir [--copy-xlsm]

# example
$ ExcelMacroExtrator.exe "C:\Dev\File.xslm" "C:\Dev\File-Source" --copy-xlsm
~~~

The `--copy-xlsm`-option copies the Excel-file to the `targetpath` as well. 

## Motivation

I develop a fair amount of Excel-VBA-Applications, and want to track the VBA-modules via `git`. The Excel-file itself is a binary-file, and therefore not really `diff`-able. With this tool I can extract the VBA-code to a `targetdir` and track that directory with `git`.
