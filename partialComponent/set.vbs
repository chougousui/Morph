' 写文件
' 1. 用户使用sheet名与range指定要写哪一格
' 2. 组件修改该处内容并回显


function configs()
    ' 必要,组件的用例
    ' Dict{"bash", "powershell"}
    set example = createObject("Scripting.Dictionary")
    with example
        .Add "bash", "morph set ./assets/test*/*.txt /val:placeholder"
        .Add "powershell", "morph set .\assets\test*\*.txt /val:placeholder"
    end with

    ' 必要,组件的操作权限
    ' true/false
    readonly = false

    ' 可选,混入的其他可独立处理过程
    ' Array(mixin1, mixin2,)
    mixins = Array("sheetMixin", "rangeMixin")

    ' 可选,参数的验证规则
    ' Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
    set validationRules = createObject("Scripting.Dictionary")
    with validationRules
        .Add "val", Array("required")                            ' 必须指定修改成什么
    end with

    set mConfig = createObject("Scripting.Dictionary")
    with mConfig
        .Add "example", example
        .Add "readonly", readonly
        .Add "mixins", mixins
        .Add "validationRules", validationRules
    end with
    set configs = mConfig
end function

' 必要,为了使realWork不做任何参数校验的前置处理
' 参数的正确性检查一般已经通过
' 此处大多数用来设置默认值
function layeredProcess(byRef xlFile, params)
    if not params.exists("sheetParam") then
        params.item("sheetParam") = 1
    end if

    set layeredProcess = params
end function

' 必要,根据参数要求,执行操作
sub realWork(byRef xlFile, params)
    ' 取参数
    loc = params.item("range")
    sheetParam = params.item("sheetParam")
    val = params.item("val")

    ' 实际处理
    set xlSheet = xlFile.worksheets(sheetParam)         ' 获取对应sheet
    xlSheet.Range(loc).Value = val                      ' 获取对应格子中的值
    wscript.echo(loc & " on sheet " & sheetParam & " set to " & val)      ' 信息回显
end sub

' 加载后的欢迎语句
wscript.echo("component: set.vbs loaded")

'-----------------------------------------------