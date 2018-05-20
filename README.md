# Morph
Unix-like VBS function shortcuts

### purpose

- To create a way to edit excel files just like plain text files, especially if the files you want to edit are scattered in different places and your computer is slow like a snail
- To extract pleasure from what is a bitter life to others(now, in Japan).

### usage

here is a **short** list for all the functions available.

- get
- grep
- incre

##### get

Get the value of a specified position in files that matched by wildcard

```powershell
cscript .\dev.vbs get .\*\*.xlsx 1 "A1"
```

- `.\dev.vbs` this tool
- `get` one of the method available
- `.\*\*.xlsx` path with wildcard
- `1` order of sheet inside excel file
- `"A1"` the specified location mentioned above

##### grep

Find all cell contents in files if meet the preset match pattern.

```powershell
cscript .\dev.vbs grep .\*\*.xlsx 1 "user-set-regex-pattern"
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

### TODO

add find-and-replace functions(the following is examples)

- replace all a to b
- replace [a,b,c,d] to [a1,b,c1,d1] if there is a [a1,b,c1,d1] in another file.