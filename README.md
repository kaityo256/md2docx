# Convert Markdown to Docx

[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

[Japanese](README_ja.md)/ English

## Summary

A sample script which converts Markdown (\*.md) to Office Open XML Document (\*.docx).
It only supports headers, bullet items, and numeric items up to 3rd depth.

Please find the explanation of how it works written in Japanese [here](http://qiita.com/kaityo256/items/7794a671d2ff8d00e603).

## Disclaimer

USE THIS SCRIPT AT YOUR OWN RISK.

## Files

* `md2docx.rb` Convert script
* `sample.md` Sample markdown file
* `template.docx` Template file

## Supporting Formats

Following items up to 3rd depth.

* header
* bullet item
* numeric item

```md:
# header1
## header2
### header3

* bullet item 1
    * bullet item 2
        * bullet item 3

1. numeric item 1
    1. numeric item 2
        1. numeric item 3
```

## Usage

```sh
$ ruby md2docx.rb
Usage: md2docx [options] file
    -t, --template [template file]
    -o, --output [output file]
```

## Results

Here is the sample input file.

```md
# md2docx sample file

## Paragraph

This is a paragraph.

## list

### Numeric item

1. hoge1
    1. hoge2.1
    1. hoge2.2
1. fuga1
    1. fuga2.1
    1. fuga2.2

### Bullet items

* bullet1
    * bullet2
        * bullet3
* bullet1
    * bullet2
        * bullet3

### Mixed

* bullet1
    1. enum1
        * bullet3
    1. enum2
```

You can convert the above via following command.

```sh
$ ruby md2docx.rb sample.md
Using template.docx
Reading sample.md
Generating sample.docx
Done.
```

If you run the above, you will have the following `sample.docx`.

![sample.png](sample.png)

It contains Japanese Katakana since the template file contains it.