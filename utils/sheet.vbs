' 获取有关sheet的参数
' para: xlWorkbook, arguments.Named
' return: 1(number) if not any sheet arguments
' return: (number) if shtord exists
' return: (string) if shtnm exists
function getSheetParam(params)
    ' 处理sheet order或sheet name,要么输出sheet name(string形式),要么输出sheet order(int形式)
    shtnm = params.item("shtnm")
    shtord = clng(orDefault(params.item("shtord"), "1"))

    ' 若已经指定sheet name,则使用name
    if len(shtnm) > 0 then
        getSheetParam = shtnm
        exit function
    end if

    ' 剩余则使用sheet order
    getSheetParam = shtord
end function

' 验证sheet有关参数
' return: false if sheet order not exists
' return: false if sheet name not exists
' return: true if in other cases
function validateSheetParam(xlFile, sheetParam)
    ' https://docs.microsoft.com/en-us/previous-versions/visualstudio/aa263402(v=vs.60)?redirectedfrom=MSDN
    ' const vbInteger = 2
    ' const vbLong = 3
    ' const vbString = 8
    if VarType(sheetParam) = vbInteger or VarType(sheetParam) = vbLong then
        ' 若为index,则验证是否超出范围,注意and操作符不会停止,不能写成一行if...
        if sheetParam > xlFile.worksheets.Count then
            wscript.echo("sheet order " & shtord & " not exists, skip")
            validateSheetParam = false
            exit function
        end if
    else
        ' 若为sheet名,目前官方没有提供worksheets.exists()方法,API很不优雅,只能手动实现
        ' 鉴于手动实现有: 循环法,试错法,外部依赖法,可能会改动,提取出函数
        if not sheetNameExists(xlFile, sheetParam) then
            wscript.echo("sheet name " & sheetParam & " not exists, skip")
            validateSheetParam = false
            exit function
        end if
    end if
    validateSheetParam = true
end function

' 判断sheet名是否存在
' return: true/false
' TODO 遍历法,试错法,isRef法
function sheetNameExists(xlFile, sheetName)
    for each sheet in xlFile.worksheets
        if sheet.name = sheetName then
            sheetNameExists = true
            exit function
        end if
    next
    sheetNameExists = false
end function