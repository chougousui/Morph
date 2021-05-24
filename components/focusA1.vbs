' 聚焦A1
' 1. 无需参数
' 2. 每一个sheet聚焦到A1
' 3. 整个文档聚焦到第一个非隐藏的sheet

' 必要,组件的使用用例
' return: Dict{"bash", "powershell"}
function example()
    set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "bash", "morph focusA1 ./assets/test*/*.xlsx"
        .Add "powershell", "morph focusA1 .\assets\test*\*.xlsx"
    end with
    set example = mDict
end function

' 必要,组件的操作权限
' return: boolean
function readonly()
    ' 更改focus需要保存
    readonly = false
end function

' 必要,morph调用的组件主操作,包含验证过程
sub extension(byRef xlFile, params)
    ' focusA1没有动态参数验证

    call afterCheck(xlFile, params)
end sub

' 单独测试组件时会使用的,不带参数校验的版本
sub afterCheck(byRef xlFile, params)
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