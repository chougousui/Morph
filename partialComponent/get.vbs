' 读文件
' 1. 用户使用sheet名与range指定要读哪一格
' 2. 组件打印出该处内容

function configs()
    ' 必要,组件的用例
    ' Dict{"bash", "powershell"}
    set example = createObject("Scripting.Dictionary")
    with example
        .Add "bash", "morph get ./assets/test*/*.txt"
        .Add "powershell", "morph get .\assets\test*\*.txt"
    end with

    ' 必要,组件的操作权限
    ' true/false
    readonly = true

    ' 可选,混入的其他可独立处理过程
    ' Array(mixin1, mixin2,)
    mixins = Array("sheetMixin", "rangeMixin")

    set mConfig = createObject("Scripting.Dictionary")
    with mConfig
        .Add "example", example
        .Add "readonly", readonly
        .Add "mixins", mixins
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

    ' 实际处理
    set xlSheet = xlFile.worksheets(sheetParam)         ' 获取对应sheet
    res = xlSheet.Range(loc).Value                      ' 获取对应格子中的值
    wscript.echo(res)                                   ' 打印至标准输出
end sub

' 加载后的欢迎语句
wscript.echo("component: get.vbs loaded")


' ------------------------------------------------------------------------