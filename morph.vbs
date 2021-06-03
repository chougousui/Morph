' 提供常用的基础功能
' 1. 解析路径
' 2. 加载外部文件
' 3. 字符串到函数(以验证规则取验证函数会用到)
' https://docs.microsoft.com/en-us/previous-versions//xe43cc8d(v=vs.85)

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
    ' wscript.echo(fSpec)
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

' 程序的入口文件放在 entry.vbs 中的main函数中
Include(resolvePath("./entry.vbs"))
main()