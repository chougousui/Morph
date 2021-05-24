' 写文件
' 1. 用户使用sheet名与range指定要写哪一格
' 2. 组件修改该处内容并回显

' 必要,组件的使用用例
' return: Dict{"bash", "powershell"}
function example()
    set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "bash", "morph set ./assets/test*/*.txt [/shtord:1 | /shtnm:""Sheet1""] /range:B1 /val:placeholder"
        .Add "powershell", "morph set .\assets\test*\*.txt [/shtord:1 | /shtnm:""Sheet1""] /range:B1 /val:placeholder"
    end with
    set example = mDict
end function

' 必要,组件的操作权限
' return: boolean
function readonly()
    ' 操作结果需要保存
    readonly = false
end function

' 可选,参数的验证规则
' return Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
function validationRules()
    ' 返回该插件要求的验证规则
    Set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "shtord", Array("nullable", "isInt", "positive")   ' 如果sheet order存在,则必须是正整数
        .Add "shtnm", Array("nullable")                         ' sheet name
        .Add "range", Array("required", "isSingleRef")          ' 需要指定修改位置
        .Add "val", Array("required")                           ' 需要指定修改成什么
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
function extension(xlFile, params)
    ' 加载参数默认值工具
    Include(resolvePath("./utils/defaultValue.vbs"))
    ' 加载sheet相关工具
    Include(resolvePath("./utils/sheet.vbs"))

    ' 获取sheet相关参数
    sheetParam = getSheetParam(params)

    ' 校验sheet相关参数
    valid = validateSheetParam(xlFile, sheetParam)
    if not valid then
        exit function
    end if

    ' 检查完毕后开始使用组件
    call afterCheck(xlFile, params, sheetParam)
end function

' 单独测试组件时会使用的,不带参数校验的版本
function afterCheck(byRef xlFile, params, sheetParam)
    loc = params.item("range")
    val = params.item("val")

    ' 获取sheet
    set xlSheet = xlFile.worksheets(sheetParam)

    ' 操作sheet值
    xlSheet.Range(loc).Value = val

    ' 信息回显
    wscript.echo(loc & " on sheet " & sheetParam & " set to " & val)
end function

' 加载后的欢迎语句
wscript.echo("component: set.vbs loaded")