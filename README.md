# word-macos-bookmarks-fix

This repository provides a small VBA and Python script so that, when exporting a PDF using Word for macOS without using Microsoft's Web Services (or when exporting a PDF using the "Print" function), it's still possible to have embedded bookmarks/outlines and file metadata.

## Installation

- First, you'll need to import (just copy and paste) the [WordMarco.bas](./WordMacro.bas) file on Word. Enable the Developer options ribbon from Word's settings, then click on the `Macro` option, write the name of the new macro, click on the "+" and paste in the editor the code.
- Then, you'll need to download the [AddTitlesToPdf.py](./AddTitlesToPdf.py) Python script. It requires the `pypdf` library, so you'll need to manually install it using `pip install pypdf`.

## Usage

First, run the Macro on Word. It'll generate a text file, save it where you want. Now, run the AddTitlesToPdf Python script: `python3 AddTitlesToPdf.py *PDF File* *Text File*`. The script will add the bookmarks and the metadata to the PDF file.

### Settings

| Command | Description | Followed by |
| - | - | - |
| `--skip-metadata` | PDF Metadata won't be changed. | Nothing |
| `--top-points` | The bookmark link will be higher of *this* points compared to the text position. | A number |
| `--left-points` | The bookmark link will be more at the left of *this* points compared to the text position. | A number |
| `--put-wrong-outline-first` | Sometimes, the Word file might have a outline/bookmark structure that it's impossible to copy in a PDF (for example, on Word you can put a bookmark of level 7 below a bookmark of level 2, but this isn't permitted in PDFs). By defualt, the script will decrease the bookmark level until it's possible (in the example above, the bookmark will become of level 3), but, if you pass this argument, the script will always put it at level 1. | Nothing | 
| `--extract` | If you've exported only some of the pages of your Word document, you can specify which pages you've exported so that the script will map the bookmarks correctly. | A string that can be composed on a comma-separated list of numbers (ex: `1,2,3`); a range of pages (ex: `1-5` --> from page 1 to page 5); or a combination of these two. |
| `--start-extract-count-from` | If you've customized Word's page count, and you've generated an extract, you can put here the custom numeration (so, how many pages are in the Word document before the start of the page count) so that the script will correctly match the PDF bookmarks. | A number |
| `--output-name` | Change the output file name/path. | A string |