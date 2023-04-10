# ExcelMerger
A program that merges 2 excel files together
It only looks at the first 2 colums (A and B) and merges multiple files togetter so that missing data is filled up

## Usage
```
Usage: ExcelMerger [--output-file <String>] [--nodir] [--noprompts] [--noopen-output] [--noheader] [--noadjust-to-content] [--show-internal-errors] [--help] [--version] input-files

ExcelMerger

Arguments:
  0: input-files    Path to the directory filled with *.xlsx that need to be merged (Required)

Options:
  --output-file <String>    Path+name of the output file (Default: ExcelFileMerger_output.xlsx)
  --nodir                   If set threat the input-files argument as a list of file paths separated by ';'
  --noprompts               Answers yes to all prompts
  --noopen-output           If set don't open the output file when done processing
  --noheader                If set remove the line at the top of the excel document with the names of all the files
  --noadjust-to-content     If set don't adjust the with of the columns to the content size
  --show-internal-errors    If set show the internal errors
  -h, --help                Show help message
  --version                 Show version
 ```
