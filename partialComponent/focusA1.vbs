' 聚焦A1
' 1. 无需参数
' 2. 每一个sheet聚焦到A1
' 3. 整个文档聚焦到第一个非隐藏的sheet

function configs()
    ' 必要,组件的用例
    ' Dict{"bash", "powershell"}
    set example = createObject("Scripting.Dictionary")
    with example
        .Add "bash", "morph focusA1 ./assets/test*/*.xlsx"
        .Add "powershell", "morph focusA1 .\assets\test*\*.xlsx"
    end with

    ' 必要,组件的操作权限
    ' true/false
    readonly = false

    set mConfig = createObject("Scripting.Dictionary")
    with mConfig
        .Add "example", example
        .Add "readonly", readonly
    end with
    set configs = mConfig
end function

' 必要
function layeredProcess(byRef xlFile, byRef params)
    ' 但没有内容
    set layeredProcess = params
end function

' 必要
' para: xlWorkbook, arguments.Named
sub realWork(byRef xlFile, byRef params)
    ' 每个sheet都需要先activate再操作focus到A1
    for each sheet in xlFile.worksheets
        if sheet.visible then
            sheet.Activate
            sheet.Range("A1").Activate
            xlFile.windows(1).scrollRow = 1
            xlFile.windows(1).scrollColumn = 1
        end if
    next

    ' 整个文档也跳转到第一个非隐藏的sheet
    for each sheet in xlFile.worksheets
        if sheet.visible then
            sheet.Activate
            exit for
        end if
    next

    ' 没有内容修改时必须手动调用保存才能保存
    xlFile.save
    wscript.echo("focuses saved")
end sub

' 加载后的欢迎语句
wscript.echo("component: grep.vbs loaded")


'----------------------------------