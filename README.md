# Opj file Exporter
This tool is extracting *.OPJ* binary file data in text (_.txt_) file.

## Logic
The algorith is looping over excels created in the Origin file.
And then, over the spreadsheets contained in the excel.

> Only, the **data** spreadsheet is extracted.

## Usage
The usage is pretty simple :
`./exportOpjFile "path/to/file"`

The output generated is a number of files (as many as excels in the file), inside a folder.
The folder is at the path that the parsed file under the same name of the parsed file (without the file extension).

### Dependencies
The executable is based on the __**liborigin**__ [library](https://sourceforge.net/projects/liborigin/).

The version used is the _actual_ latest (**3.0.0**).

This tool, hence, inherits all the limitations that the library adds.

### Downloads
Build folders are available for each architecture (32 and 64bits).

> So, download the _one_ appropriate to your own computer architecture.

#### Contributions
Feel free to contribute by forking or push requesting.
Don't hesitate to submit an issue if needed.

#### Cheers
Amaury
