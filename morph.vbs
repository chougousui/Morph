' 主文件
' 包含许多基础组件(解析路径,加载外部文件)
' 处理命令的相关操作(参数解析和显示, 抛出错误等)
' https://docs.microsoft.com/en-us/previous-versions//xe43cc8d(v=vs.85)

' 入口函数
sub main()
    dim subCommand     ' 子命令
    dim wildcard       ' 操作对象路径描述

    dim options        ' 具名参数列表,用于传递给子组件
    dim files          ' 解析通配符后的文件名列表

    ' 只在主脚本中使用无命名参数
    ' 用例: morph get ./assets/*.xls /range:A1
    if wscript.Arguments.Unnamed.Count < 2 then
        msg = "invalid base arguments" & vbcrlf _
            & vbcrlf _
            & helpMessage()

        call err.raise(5, "", msg)
    end if

    ' 获取必要参数
    subCommand = wscript.Arguments.Unnamed(0)
    wildcard = wscript.Arguments.Unnamed(1)
    set options = wscript.Arguments.Named

    ' 根据子命令加载同名组件
    pluginPath = "./components/" & subCommand & ".vbs"
    Include(resolvePath(pluginPath))

    ' 加载验证函数
    Include(resolvePath("./validators.vbs"))

    ' 加载组件包装器
    Include(resolvePath("./componentWrapper.vbs"))

    ' 使用加载到的验证函数,静态验证组件使用的具名参数
    call wrappedValidation(subCommand, options)

    ' 加载文件操作函数
    Include(resolvePath("./fileOperator.vbs"))

    ' 一切验证就绪后开始解析操作文件,而后打开app操作
    set files = resolvePaths(wildcard)
    call processFiles(files, options)
end sub

' 显示本工具的帮助信息
' return: string
function helpMessage()
    availableSubCommands = Array("eg", "focusA1", "get", "set", "grep")

    helpMessage = "usage:" & vbcrlf _
                & "morph <subCommand> <wildcard> [options]" & vbcrlf

    helpMessage = helpMessage _
                & vbcrlf _
                & "available subCommand:" & vbcrlf

    for each command in availableSubCommands
        helpMessage = helpMessage & "- " & command & vbcrlf
    next

    helpMessage = helpMessage _
                & vbcrlf _
                & "e.g.(bash):" & vbcrlf _
                & "morph eg . /help:" & vbcrlf _
                & "e.g.(powershell):" & vbcrlf _
                & "morph eg . /help"
end function

' 调用powershell功能来解析通配符路径
' return: List{path1, path2, }
function resolvePaths(wildcard)
    ' 依靠powershell的功能,将各种格式的glob path转换为完整路径的文件列表
    ' resolve-path 接通配符时不会报错
    set res = createobject("system.collections.arraylist")

    strPS1Cmd = "(resolve-path " & wildcard & ").path"
    set objShell = wscript.createobject("wscript.shell")
    set oExec = objShell.Exec("powershell -command """ & strPS1Cmd & """ ")
    set oStdOut = oExec.StdOut

    do while not oStdOut.AtEndOfStream
        path = oStdOut.ReadLine
        res.add path
    Loop

    set resolvePaths = res
end function

' 调用powershell功能来解析通配符路径,可报错
' return: path
' return: error
function resolvePath(relative)
    ' 依靠powershell的功能,将各种格式的glob path转换为完整路径的文件
    ' resolve-path 接相对路径时,能够报错说找不到
    strPS1Cmd = "(resolve-path " & relative & ").path"
    set objShell = wscript.createobject("wscript.shell")
    set oExec = objShell.Exec("powershell -command """ & strPS1Cmd & """ ")
    set oStdOut = oExec.StdOut
    set oStdErr = oExec.StdErr

    if not oStdErr.AtEndOfStream then
        firstTwoLines = oStdErr.ReadLine & oStdErr.ReadLine
        call err.raise(432, "", firstTwoLines)
    end if

    resolvePath = oStdOut.ReadLine
end function

' 引用外部文件
' return: error
Sub Include(fSpec)
    ' 加载外部文件
    ' 发生在打开文件之前,直接抛出错误即可,无需清理资源
    With CreateObject("Scripting.FileSystemObject")
        if (.fileExists(fSpec)) then
            executeGlobal .openTextFile(fSpec).readAll()
        else
            call err.raise(432, "", "file: [" & fSpec & "] not exists! cannot include")
        end if
    End With
End Sub

' 自定义vbs的基础反射功能,从字符串到函数
' return: function, if found
' return: nothing
' side effect: print message
function myGetRef(target, needPrint)
    ' 只因getref的报错太过模糊
    ' 在遇到报错是返回nothing,让外部决定具体报什么错误

    on error resume next
    set f = getRef(target)
    if err.number = 0 then
        set myGetRef = f
        exit function
    elseif err.number = 5 and needPrint then
        wscript.echo("function " & target & " not defined")
    end if

    set myGetRef = nothing
end function



'---------------------------------
main()