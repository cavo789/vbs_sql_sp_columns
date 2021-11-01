# SQL Server - Document tables, extract structure as CSV file

![Banner](./banner.svg)

> Connect to SQL Server, iterate every tables and generate a .csv file by table containing the table's structure

## Table of Contents

- [Description](#description)
- [Install](#install)
- [Usage](#usage)
- [Author](#author)
- [License](#license)

## Description

Connect to a SQL Server database, obtain the list of
tables in that db (process all schemas), get the structure
of each tables thanks the sp_columns stored procedure and
for each table, export that structure in a results subfolder

At the end, we'll have as many files as there are tables in
the database. One .CSV file by table.

The content of the CSV will be what is returned by the sp_columns
stored procedure.

## Install

Get a copy of the script, save it to your computer.
Get also a copy of the `test.bat` file and edit that file.

See below, you'll need to mention four parameters

```
cscript.exe sql_sp_columns.vbs "servername" "dbname" "login" "password"
```

Note that you can also edit the first lines of the `sql_sp_columns.vbs` file and mention these infos immediatly as constants.

## Usage

Just run the `batch` file from a command prompt.

A connection to your SQL DB will be made and one .csv file and one .md file will be created in the `results` subfolder; one by table.

## Author

AVONTURE Christophe

## Contribute

PRs not accepted.

## License

[MIT](LICENSE)
