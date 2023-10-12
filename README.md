# ARB Excel

For reading, creating and updating ARB files from XLSX files.

## Install
1. Download repo
2. Use command below:
```bash
dart pub global activate --source path <path>
```
<Path> is your local path to the repo.

## Usage

From ARB to Excel:
```bash
dart pub global run arb_excel_dor -e <file_name1.arb> <file_name2.arb> …
```

It will create excel file with your translations.
It is recommended to remove language suffix from file name.

From Excel to ARB:
```bash
dart pub global run arb_excel_dor -a <excel_file_name.xlsx>
```
It is necessary to use that method while your directory is where the file is.
