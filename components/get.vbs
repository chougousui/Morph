' 读文件
' 1. 用户使用sheet名与range指定要读哪一格
' 2. 组件打印出该处内容

' 必要,组件的使用用例
' return: Dict{"bash", "powershell"}
function example()
    set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "bash", "morph get ./assets/test*/*.txt [/shtord:1 | /shtnm:""Sheet1""] /range:A1"
        .Add "powershell", "morph get .\assets\test*\*.txt [/shtord:1 | /shtnm:""Sheet1""] /range:A1"
    end with
    set example = mDict
end function

' 必要,组件的操作权限
' return: boolean
function readonly()
    ' 只是get,所以只读
    readonly = true
end function

' 可选,参数的验证规则
' return Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
function validationRules()
    ' 为使组件简单,以数组形式返回,由外部代码做操作
    Set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "shtord", Array("nullable", "isInt", "positive")   ' 如果sheet order存在,则必须是正整数
        .Add "shtnm", Array("nullable")                         ' sheet name
        .Add "range", Array("required", "isSingleRef")          ' 要求必须指定读哪里
    end with
    set validationRules = mDict
end function

'可选,组件的跨字段验证规则
' return Dict{{rule1, (opt1, opt2)},{rule2, (opt1, opt2)},}
function crossValidationRules()
    set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "conflict", Array("shtord", "shtnm")               ' sheet order和sheet name不能同时存在
    end with
    set crossValidationRules = mDict
end function

' 必要,morph调用的组件主操作,包含验证过程
sub extension(byRef xlFile, params)
    ' 加载参数默认值工具
    Include(resolvePath("./utils/defaultValue.vbs"))
    ' 加载sheet相关工具
    Include(resolvePath("./utils/sheet.vbs"))

    ' 获取sheet相关参数
    sheetParam = getSheetParam(params)

    ' 校验sheet相关参数
    valid = validateSheetParam(xlFile, sheetParam)
    if not valid then
        exit sub
    end if

    ' 校验其他参数
    ' 那些只有打开文件后才能做的校验

    ' 校验完毕后开始使用组件
    call afterCheck(xlFile, params, sheetParam)
end sub

' 单独测试组件时会使用的,不带参数校验的版本
sub afterCheck(byRef xlFile, params, sheetParam)
    ' 获取sheet以外的参数
    loc = params.item("range")

    ' 获取对应sheet
    set xlSheet = xlFile.worksheets(sheetParam)

    ' 而后读取sheet中位于loc格中的内容
    res = xlSheet.Range(loc).Value

    ' 打印至标准输出
    wscript.echo(res)
end sub

' 加载后的欢迎语句
wscript.echo("component: get.vbs loaded")

' 手动单独调试时的入口
sub manualMain()
    set xlApp = createobject("Excel.Application")
    file = "C:\Users\XXX\morph\assets\test1\1.xlsx"
    set xlFile = xlApp.workbooks.open(file, , true)

    call afterCheck(xlFile, wscript.Arguments.Named, "Sheet1")

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