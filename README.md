# TransExc2Csv
Batch-Convert File Tool

Tool intent:
     Achieve work file extraction of multiple Sheet tab target data of Excel, and output CSV file.

Function Description:
     The tool implements Excel data conversion in CSV format.
     1. Convert all .xls / xlsx files in the specified directory;
     2. Corresponding to multiple Excel tabs, they are stored in the same directory level, and the difference is named "(sheet)".
Code structure:
     Tool entry file: happy_transform.py
     Parameters: path-file path (direct directories can be used)

User guides:
     1.win + r-> cmd
     2. python happy_transform.py "D: \ Data \ test \ XXX.xls"
         or
        python happy_transform.py "D: \ Data \ test"
