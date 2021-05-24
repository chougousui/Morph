' 放置excel文件操作相关的所有代码
' 文件打开
' 文件读取
' 文件关闭

' 全局变量,为方便无参数调用handleError
dim xlApp        ' excel app
dim xlFile       ' excel workbook

' 主函数
sub processFiles(files, options)
    ' 打开app,此处之后的任意错误都应该catch住,并在清理后结束程序
    on error resume next
    set xlApp = createobject("Excel.Application")

    for each file in files
        ' 先显示操作对象
        call wrappedFilePrompt(file)

        ' 打开文件并操作
        isReadOnly = readonly()
        set xlFile = xlApp.workbooks.open(file, , isReadOnly)
        call wrappedExtension(xlFile, options)

        ' 及时错误处理,不要让保存文件得以运行
        if err.number <> 0 then
            wscript.echo("line 26")
            handleError()
        end if

        ' 处理完一个就关闭一个
        call closeCurrentFile(xlFile)

        ' 及时处理错误,不让下一个for loop的文件被打开
        if err.number <> 0 then
            wscript.echo("line 35")
            handleError()
        end if
    next

    if err.number <> 0 then
        wscript.echo("line 41")
        handleError()
    end if

    clearBeforeQuit()

    if err.number = 0 then
        wscript.echo(vbcrlf & "all files specified processed successfully")
    end if
end sub

' 错误处理,遇到错误会清理并退出脚本运行
sub handleError()
    ' 在遇到错误时,打印错误信息,并立即退出程序处理
    ' 然而VBS只能打印非常有限的信息
    wscript.echo("Source = " & err.Source & ", # = " & err.Number &", Desc=" & err.Description)

    ' 清理环境
    clearBeforeQuit()

    wscript.quit()
end sub

' 退出前尝试关闭打开的文件和app
sub clearBeforeQuit()
    On Error Resume Next

    call closeCurrentFile(xlFile)

    call quitApp(xlApp)
end sub

' 若有空闲未关闭的xlFile,则关闭之并提示,若没有则不做任何事
sub closeCurrentFile(byRef xlFile)
    if VarType(xlFile) = vbObject then
        if not xlFile is nothing then
            needSave = wrappedNeedSave()
            ' true表示保存内容,以只读模式打开时,会弹出页面提示
            xlFile.close(needSave)
            set xlFile = nothing
        end if
    end if
end sub

' 若有空闲未退出的xlApp,则退出之并提示,若没有则不做任何事
sub quitApp(byRef xlApp)
    if VarType(xlApp) = vbObject then
        if not xlApp is nothing then
            xlApp.Quit
            wscript.echo("app quitted")
            set xlApp = nothing
        end if
    end if
end sub