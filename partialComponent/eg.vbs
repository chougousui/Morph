' 举例用组件
' 1. 可显示调用时所用参数
' 2. 可显示文件的一些sheet信息
' 3. 显示一句话

function configs()
    ' 必要,组件的用例
    ' Dict{"bash", "powershell"}
    set example = createObject("Scripting.Dictionary")
    with example
        .Add "bash", "morph eg ./assets/test*/*.xlsx /testInt:1 /testString:A"
        .Add "powershell", "morph eg .\assets\test*\*.xlsx /testInt:1 /testString:A"
    end with

    ' 必要,组件的操作权限
    ' true/false
    readonly = true

    ' 可选,用于混入其他独立处理过程
    ' Array(mixin1, mixin2,)
    mixins = Array("sheetMixin", "rangeMixin")
    ' 1. sheetMixin 添加选项卡shtord, shtnm,及其相关验证,输出sheetParam
    ' 2. rangeMixin 添加选项卡range及其相关验证,输出range
    ' vbs无法在多行定义Array的同时插入注释

    ' 可选,参数的验证规则
    ' Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
    set validationRules = createObject("Scripting.Dictionary")
    with validationRules
        .Add "testInt", Array("nullable", "isInt", "positive")      ' 如果有,必须是正整数
        .Add "testString", Array("nullable")
        .Add "shtord", Array("required")                            ' 在mixin的基础上添加验证
    end with

    ' 可选,参数的交叉验证规则
    set crossValidationRules = CreateObject("System.Collections.ArrayList")
    with crossValidationRules
        .Add Array("requiredAny", Array("testInt", "testString"))   ' testInt和testString虽然每个都是nullable,但至少指定一个
    end with

    set mConfig = createObject("Scripting.Dictionary")
    with mConfig
        .Add "example", example
        .Add "readonly", readonly
        .Add "mixins", mixins
        .Add "validationRules", validationRules
        .Add "crossValidationRules", crossValidationRules
    end with
    set configs = mConfig
end function

' 必要
' 为了使realWork不做任何参数校验的前置处理
' para: xlWorkbook, arguments.Named
function layeredProcess(byRef xlFile, byRef params)
    ' 可在找不到参数时报错,不过通常这种校验已经做过
    if not params.exists("range") then
        call err.raise(5, "", "range is must in [eg]")
    end if

    ' 常用来设置参数默认值
    if not params.exists("sheetParam") then
        params.item("sheetParam") = 1
    end if

    ' 还可打印参数,用于校验morph对于参数的处理结果
    wscript.echo("options:")
    wscript.echo("Key" & vbtab & vbtab & "Value")
    for each key in params
        value = params.item(key)
        wscript.echo(key & vbtab & vbtab & value)
    next

    ' 还可读取一些文件信息,帮助后续操作选择合理的地址范围
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

    set layeredProcess = params
end function

' 单独测试组件时会使用的,不带参数校验的版本
' para: xlWorkbook, arguments.Named
sub realWork(byRef xlFile, byRef params)
    wscript.echo(vbcrlf & "component function works in every file")
end sub

' 加载后的欢迎语句
wscript.echo("component: eg.vbs loaded")

' ------------------------------------------------------------------------
' 手动单独调试时的入口
sub manualMain()
    set xlApp = createobject("Excel.Application")
    file = "C:\Users\XXX\morph\assets\test1\1.xlsx"
    set xlFile = xlApp.workbooks.open(file, , true)

    call realWork(xlFile, wscript.Arguments.Named)

    xlFile.close(false)
    set xlFile = nothing
    xlApp.quit
    set xlApp = nothing
end sub

' 手动加载,单独调试
if wscript.arguments.Unnamed.Count = 0 then
    manualMain()
end if