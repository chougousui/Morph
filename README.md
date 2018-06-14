# Morph
Unix-like VBS function shortcuts

### Purpose

- To create a way to edit excel files just like plain text files, especially if the files you want to edit are scattered in different places and your computer is slow like a snail
- To extract pleasure from what is a bitter life to others(now, in Japan).

### Usage

Here is a **short** list for all the functions available.

- get
- set
- grep
- find
- findComment
- incre
- replace

### Basic functions

Three parameters are required to accomplish the basic functions

```powershell
cscript .\dev.vbs <operation> <globPath> <sheetOrder>
```

- This will execute the chosen operation on each matched files
- Wildcards are supported here to match file names
  - but `git-bash` will automatically parse the globPath and pass the first of the matched files to the program. The right way to pass globPath to script is to use `""` around the globPath, no matter what kind of style the path string is.
  - match `xls,xlsx,xlsm` files by default. (otherwise it is not necessary to use VB Script)
  - the globPath is set to be impermeable , which means you can match "./1/eg1.xls" with "./1/\*", but can not match "./1/1/eg2.xls" with the same "./1/\*", I know this is counter-intuitive, but you need to know that someone likes to place multiple backup folders in the same folder. Handling files under these subfolders is meaningless.
- both name and order (starts from 1) can be used as locator of a sheet, and the script will print the sheet name for confirm.

### Operations

##### get

Get the value of a specified position in files

```powershell
cscript .\dev.vbs get .\*\*.xlsx 1 "A1"
```

- `.\dev.vbs` this tool
- `get` one of the method available
- `.\*\*.xlsx` path with wildcard
- `1` order of sheet inside excel file
- `"A1"` the specified location mentioned above

##### set

Set the value of a specified positon in files

```powershell
cscript .\dev.vbs set .\*\*.xlsx 1 "A1" "toString"
```

- Just like that in `get`
- File will be saved by default, so remember to backup important files

##### grep

Find all cell contents in files.

```powershell
cscript .\dev.vbs grep .\*\*.xlsx 1 "user-set-regex-pattern"
```

##### find

Find all cell contents within a specified range.

```powershell
cscript .\dev.vbs find .\*\*.xlsx 1 "values" "A:B"
```

##### findComment

Find all cell contents within a specified range, using a piece of comment.

```powershell
cscript .\dev.vbs findComment .\*\*.xlsx 1 "comment" "A:B"
```

##### incre

Finding cells that match a certain format within a certain range of conditions

```powershell
cscript .\dev.vbs incre .\*\*.xlsx 1 "pattern-to-search range" "(xstart,xend,ystart,yend)" "second-pattern-to-search-further-content"
```

eg:

```powershell
cscript .\dev.vbs incre .\*\*.xlsx 1 "list" "(0,4,0,3)" "table"
```

In the example,I'm trying to search the name of tables under a header--"list",like this:

 (Pretend you know that some Japanese people love to use excel to save design documents)

| Used TableList |               |      |      |
| -------------- | ------------- | ---- | ---- |
|                | addressTable  |      |      |
|                | nameTable     |      |      |
|                | shoppingTable |      |      |
|                |               |      |      |
|                |               |      |      |

`(0,4,0,3)`means that you want to expand the cell (matched with "list") to a range like this:

from `(cell.row + 0, cell.column + 0)`  to `(cell.row + 4, cell.column + 3)`

##### replace

Replace all `fromString` with `toString` in the file, just like that in excel(with default option)

```powershell
cscript .\dev.vbs replace .\*\*.xlsx 1 "fromString" "toString"
```

- TODO: Show a counter if possible

### Others

- If you are using git-shell or something similar instead of windows powershell, you may need to pay attention to the problem of text encoding. In order to ensure that git-shell ,which display text with utf-8 ,can display another text encoding properly, you may need to do these:

  - Convert the file to `ANSI` and select an encoding such as `GBK`  or `Shift-JIS` until you can see normal text .

    - The appropriate text encoding depends on the region.
    - notepad++ is recommended

  - Use a pipeline command to convert the output of VBS program

    ```shell 
    cscript ./dev.vbs get .\*\*.xlsx 1 "A1" | iconv -c -f Shift-JIS -t UTF-8
    ```

    - This convert the output of `dev.vbs` from `Shift-JIS` into `UTF-8` and skip the text that failed to convert by using option "-c".

- Add a non-generic feature to the dev branch for filling in other forms based on an overview file by a key

  ```powershell
  cscript ./dev.vbs fill .\test\*dstFile* 4 "D9" D:\srcFile.xlsx ALL
  ```

  - now the script will assign the result to null actively if it can't find any information about the key in overview file

### TODO


- replacement counter