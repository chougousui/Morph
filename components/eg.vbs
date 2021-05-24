' 举例用组件
' 1. 可显示调用时所用参数
' 2. 可显示文件的一些sheet信息
' 3. 显示一句话

' 必要,组件的使用用例
' return: Dict{"bash", "powershell"}
function example()
    set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "bash", "morph eg ./assets/test*/*.xlsx /range:A1 /test:1"
        .Add "powershell", "morph eg .\assets\test*\*.xlsx /range:A1 /test:1"
    end with
    set example = mDict
end function

' 必要,组件的操作权限
' return: boolean
function readonly()
    ' 用于决定该组件在操作文件时的权限
    readonly = true
end function

' 可选,参数的验证规则
' return Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
' 只返回规则,具体验证由componentWrapper完成
function validationRules()
    set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "range", Array("required", "isSingleRef")          ' 必须有
        .Add "test", Array("nullable", "isInt", "positive")     ' 如果有,必须是正整数
    end with
    set validationRules = mDict
end function

' 必要,morph调用的组件主操作,包含参数和文件信息验证过程
' para: xlWorkbook, arguments.Named
sub extension(byRef xlFile, params)
    ' 打印参数
    wscript.echo("options:")
    wscript.echo("Key" & vbtab & "Value")
    for each key in params
        value = params.item(key)
        wscript.echo(key & vbtab & value)
    next

    ' 读取一些文件信息
    wscript.echo(vbcrlf & "sheets:")
    wscript.echo("Index" & vbtab & "Name" & vbtab & "Visibility" & vbtab & "Used Range")
    for each xlSheet in xlFile.worksheets
        visibleString = "visible"
        if not xlSheet.visible then
            visibleString = "invisible"
        end if
        wscript.echo(xlSheet.index & vbtab & xlSheet.name & vbtab & visibleString _
                    & vbtab & vbtab & xlSheet.UsedRange.Address)
    next

    ' 参数信息显示/校验完毕后,开始使用组件
    call afterCheck(xlFile, params)
end sub

' 单独测试组件时会使用的,不带参数校验的版本
' para: xlWorkbook, arguments.Named
sub afterCheck(byRef xlFile, params)
    wscript.echo(vbcrlf & "component function works in every file")
end sub

' 加载后的欢迎语句
wscript.echo("component: eg.vbs loaded")

' 手动单独调试时的入口
sub manualMain()
    set xlApp = createobject("Excel.Application")
    file = "C:\Users\XXX\morph\assets\test1\1.xlsx"
    set xlFile = xlApp.workbooks.open(file, , true)

    call afterCheck(xlFile, wscript.Arguments.Named)

    xlFile.close(false)
    set xlFile = nothing
    xlApp.quit
    set xlApp = nothing
end sub

' ------------------------------------------------------------------------
' 手动加载,单独调试
if wscript.arguments.Unnamed.Count = 0 then
    manualMain()
end if