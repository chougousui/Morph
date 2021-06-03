' 保存命令相关的入口函数,逻辑主干
' 1. 子命令和路径的解析
' 2. 显示命令帮助文本

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

    set files = resolvePaths(wildcard)

    ' 使用子命令对应的组件处理文件列表
    Include(resolvePath("./component.vbs"))
    set comp = (new component)(subCommand)
    comp.processFiles(files)
end sub

' 显示本工具的帮助信息
' return: string
function helpMessage()
    availableSubCommands = Array("eg", "focusA1", "get", "set", "grep")

    helpMessage = "usage:" & vbcrlf _
                & "morph <subCommand> <wildcard> [options]" & vbcrlf _
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
                & "morph eg /help"
end function