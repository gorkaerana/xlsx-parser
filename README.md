# xlsx-parser

## Background

I track my expenses using an Excel file with .xlsx format. When I decided to synchronize said file with a few tables in an Org Mode file I found the process too tedious. Namely, in order to transform an Excel table into an org table one has to:
1. Save the table in csv format (with whitespace as separator, which makes me uncomfortable).
2. Upload the .csv file into Org Mode table via the `org-table-import` function.

Since my expense tracker contains multiple tables that I'd like synchronized in my Org Mode file, the process takes a couple of minutes. So I seized this opportunity to learn Emacs Lisp by writing a little xlsx parser.

## Goal

Write a function that, given the file path of an Excel file and a sheet name, will insert an org-table where the marker is placed.