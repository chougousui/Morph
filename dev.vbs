''' global vars
Dim nowpath
Dim xlApp,xlFile,xlSheet
Dim fileList
Dim dstFileList
Dim operation,globPath,sheetOrder
Dim searchRange
Dim grepPattern
Dim firstPattern,increPattern,offset,realOffset
Dim fromString,toString
Dim startPoint,srcFilePath,fillMethod,nameList,srcMaxRow

''' main

init()

On Error Resume Next
for each file in dstFileList
    ' echo file name and open sheet
    wscript.echo replace(file,nowpath,".")
    set xlFile = xlApp.Workbooks.Open(file)
    set xlSheet = xlFile.Worksheets(sheetOrder)
	wscript.echo xlSheet.name
    ' do some operation
    if operation = "get" then
        doGet()
    elseif operation = "set" then
        doSet()
    elseif operation = "grep" then
        doGrep()
    elseif operation = "find" then
        doFind()
    elseif operation = "findComment" then
        doFindComment()
    elseif operation = "replace" then
        doReplace()
    elseif operation = "incre" then
        doIncre()
    elseif operation = "fill" then
        doFill()
    else
        wscript.echo "no operation"
    End if
	
	' echo a null line
	wscript.echo
	
    ' close file without saving
    if operation = "set" Or operation = "replace" Or operation = "fill" then
        xlFile.Close(True)
    else
        xlFile.Close(False)
    end if
next

if Err.Number <> 0 then
    wscript.echo "an error occured: " & Err.Description
    Err.Clear
End if

On Error Goto 0
' always run this even error occured
clean()

''' /main

Sub init()
    ' check parameters, at least 3 parameters: operation,globPath,sheetOrder
    if wscript.Arguments.Count < 3 then
        wscript.echo "Missing parameters"
    end if

    ' get nowpath
    nowpath = left(wscript.scriptfullname,instrrev(wscript.scriptfullname,"\")-1)

    ' init and edit parameters
    ' required parameters
    operation = wscript.Arguments(0)
    globPath = wscript.Arguments(1)
    sheetOrder = wscript.Arguments(2)
	if IsNumeric(sheetOrder) then
		sheetOrder = CInt(sheetOrder)
	end if
    if operation = "get" then
        searchRange = wscript.Arguments(3)
    elseif operation = "set" then
        searchRange = wscript.Arguments(3)
        toString = wscript.Arguments(4)
    elseif operation = "grep" then
        grepPattern = wscript.Arguments(3)
    elseif operation = "find" then
        grepPattern = wscript.Arguments(3)
        searchRange = wscript.Arguments(4)
    elseif operation = "findComment" then
        grepPattern = wscript.Arguments(3)
        searchRange = wscript.Arguments(4)
    elseif operation = "replace" then
        fromString = wscript.Arguments(3)
        toString = wscript.Arguments(4)
    elseif operation = "incre" then
        firstPattern = wscript.Arguments(3)
        increPattern = wscript.Arguments(5)
        offset = wscript.Arguments(4)
        realOffset = replace(offset,"(","")
        realOffset = replace(realOffset,")","")
        realOffset = Split(realOffset,",")
    elseif operation = "fill" then
        startPoint = wscript.Arguments(3)
        srcFilePath = wscript.Arguments(4)
        fillMethod = wscript.Arguments(5)
    end if

    ' last, launch excel, generate a filelist contains the targer files
    Set xlApp=CreateObject("Excel.Application")
    Set fileList = CreateObject("Scripting.Dictionary")
    Set dstFileList = CreateObject("Scripting.Dictionary")

    if operation = "fill" then
        ' get nameList information from srcFile
        set srcFile = xlApp.Workbooks.Open(srcFilePath,,True)
        set xlSheet1 = srcFile.Worksheets(1)
        srcUsedRow = xlSheet1.usedRange.Rows.Count
        nameList = xlSheet1.Range(xlSheet1.Cells(6,2),xlSheet1.Cells(srcUsedRow,8)).value
        srcFile.close(False)
        set xlSheet1 = Nothing
        set srcFile = Nothing
        for rowNum = UBound(nameList,1) to 1 step -1
            if not nameList(rowNum,2) = "" then
                wscript.echo nameList(rowNum,2)
                srcMaxRow = rowNum
                exit for
            end if
        next
        wscript.echo srcMaxRow
    end if

    matchFiles()
End Sub
Sub clean()
    Set xlSheet = Nothing
    Set xlFile = Nothing

    ' very important
    xlApp.Quit
    Set xlApp = Nothing
    wscript.echo
    wscript.echo "ALL CLEANED!"
End Sub
Sub doGet()
    wscript.echo xlSheet.Range(searchRange).Value
End Sub
Sub doSet()
    xlSheet.Range(searchRange).Value = toString
End Sub
Sub doGrep()
    ' find all cells
    dim rng
    set rng = xlSheet.Cells.find(grepPattern)
    if not rng is nothing then
        firstAddress = rng.Address
        Do
            wscript.echo rng.value
            set rng = xlSheet.Cells.findNext(rng)
        Loop While rng.Address <> firstAddress
    else
        wscript.echo "can not find """ & grepPattern & """"
    End If
End Sub
Sub doFind()
    ' find all cells
    dim rng
    set rng = xlSheet.Range(searchRange).find(grepPattern)
    if not rng is nothing then
        firstAddress = rng.Address
        Do
            wscript.echo rng.value
            set rng = xlSheet.Range(searchRange).findNext(rng)
        Loop While rng.Address <> firstAddress
    else
        wscript.echo "can not find """ & grepPattern & """"
    End If
End Sub
Sub doFindComment()
    ' find all cells
    dim rng

    ' -4144 means xlComment,
    set rng = xlSheet.Range(searchRange).find(grepPattern,,-4144)
    if not rng is nothing then
        firstAddress = rng.Address
        Do
            wscript.echo rng.value
            set rng = xlSheet.Range(searchRange).findNext(rng)
        Loop While rng.Address <> firstAddress
    else
        wscript.echo "can not find """ & grepPattern & """"
    End If
End Sub
Sub doReplace()
    ' replace all cells
    xlSheet.Cells.replace fromString,toString
End Sub
Sub doIncre()
    x1 = realOffset(0)
    x2 = realOffset(1)
    y1 = realOffset(2)
    y2 = realOffset(3)
    dim allRng
    dim increRng
    set allRng = xlSheet.Cells.find(firstPattern)
    if not allRng is nothing then
        metaAddress = allRng.Address
        metaCellRow = allRng.row
        metaCellCol = allRng.column
        wscript.echo "find """ & firstPattern & """ in " & metaAddress
        set newRange = xlSheet.Range(xlSheet.Cells(metaCellRow+x1,metaCellCol+y1),xlSheet.Cells(metaCellRow+x2,metaCellCol+y2))
        wscript.echo "and the new range is : " & newRange.address

        ' searching for increPattern,but the result depends on offset heavily
        set increRng = newRange.find(increPattern)
        if not increRng is nothing then
            firstAddress = increRng.Address
            Do
                wscript.echo increRng.value
                set increRng = newRange.findNext(increRng)
            Loop While increRng.Address <> firstAddress
        else
            wscript.echo "can not find """ & increPattern & """"
        End If
    else
        wscript.echo "can not find """ & firstPattern & """"
    end if
End Sub
Sub doFill()
    ' exmple:  cscript ./dev.vbs fill .\test\*画面設計書* 4 "D9" D:\＠項目_完全\＠各種一覧\E_項 目一覧（営業事務システム）.xlsx ALL
    Dim srcUsedRow,targetUsedRow,targetMaxRow
    Dim targetRange
    targetUsedRow = xlSheet.usedRange.Rows.Count
    set startCell = xlSheet.Range(startPoint)
    startRow = startCell.Row
    startCol = startCell.Column

    ' fill all
    if fillMethod = "ALL" then
        ' get targetRange,and max row
        targetRange = xlSheet.Range(xlSheet.Cells(startRow,startCol-1),xlSheet.Cells(targetUsedRow,startCol+6)).value
        for rowNum = UBound(targetRange,1) to 1 step -1
            if not targetRange(rowNum,2) = "" then
                wscript.echo targetRange(rowNum,2)
                targetMaxRow = rowNum
                exit for
            end if
        next
        wscript.echo targetMaxRow

        ' edit file
        for i=1 to targetMaxRow
            for j=1 to srcMaxRow
                if Trim(nameList(j,2)) = Trim(targetRange(i,2)) then
                    targetRange(i,1) = nameList(j,1)
                    targetRange(i,5) = nameList(j,5)
                    targetRange(i,7) = nameList(j,6)
                    targetRange(i,8) = nameList(j,7)
                    exit for
				else
					targetRange(i,1) = ""
                end if
            next
        next

        ' assign value
        xlSheet.Range(xlSheet.Cells(startRow,startCol-1),xlSheet.Cells(targetUsedRow,startCol+6)).value = targetRange

    ' fill id only
    elseif fillMethod = "ID" then
        ' get targetRange,and max row
        targetRange = xlSheet.Range(xlSheet.Cells(startRow,startCol-1),xlSheet.Cells(targetUsedRow,startCol)).value
        for rowNum = UBound(targetRange,1) to 1 step -1
            if not targetRange(rowNum,2) = "" then
                wscript.echo targetRange(rowNum,2)
                targetMaxRow = rowNum
                exit for
            end if
        next
        wscript.echo targetMaxRow

        ' edit file
        for i=1 to targetMaxRow
            for j=1 to srcMaxRow
                if Trim(nameList(j,2))=Trim(targetRange(i,2)) then
                    targetRange(i,1) = nameList(j,1)
                    exit for
                end if
            next
        next

        ' assign value
        xlSheet.Range(xlSheet.Cells(startRow,startCol-1),xlSheet.Cells(targetUsedRow,startCol)).value = targetRange
    end if
End Sub

' tools
Sub matchFiles()
    ' convert wildcards to regular expressions (only for * here)
    ' turn unix-style path to windows-style path(just in the form of a string)
    if InStr(globPath,"/") then
        namePattern = replace(globPath,"/","\\")
        namePattern = replace(namePattern,".","\.")
        namePattern = replace(namePattern,"*","[^\\]*")
    else
        namePattern = replace(globPath,"\","\\")
        namePattern = replace(namePattern,".","\.")
        namePattern = replace(namePattern,"*","[^\\]*")
    end if
	namePattern = namePattern & "xls"
	
    ' get all files under nowpath
    filetree(".")

    ' if matched, add into dstFileList
    Set fileNameRegExp = New RegExp
    for each file in fileList
        tmpfileName = replace(file,nowpath,".")
        fileNameRegExp.Pattern = namePattern
        if fileNameRegExp.Test(tmpfileName) then
            dstFileList.add file,""
        end if
    next
End Sub
Sub filetree(rootPath)
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = myFSO.getFolder(rootPath)
    set subFolders = rootFolder.SubFolders
    set files = rootFolder.Files

    ' add file under this folder into fileList
    for each file in files
        fileList.add file.path,""
    next

    ' do the same things to subFolders
    for each subFolder in subFolders
        filetree(subFolder.path)
    next
End Sub
