' 搜索文本
' 1. 在excel中搜索想要的文本
' 2. 目前支持通配符搜索
' 3. 能够指定sheet和range缩小搜索范围
' 4. 使用开关决定是否转去搜索comment

' 必要,组件的使用用例
' return: Dict{"bash", "powershell"}
function example()
    set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "bash", "morph grep ./assets/test*/*.txt /pattern:""a?c"" [/shtord:1 | /shtnm:""Sheet1""] [/range:""A1:F10""] [/comment:]"
        .Add "powershell", "morph grep .\assets\test*\*.txt /pattern:""a?c"" [/shtord:1 | /shtnm:""Sheet1""] [/range:""A1:F10""] [/comment]"
    end with
    set example = mDict
end function

' 必要,组件的操作权限
' return: boolean
function readonly()
    ' 只是grep,所以只读
    readonly = true
end function

' 可选,参数的验证规则
' return Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
function validationRules()
    Set mDict = createObject("Scripting.Dictionary")
    with mDict
        .Add "pattern", Array("required")                       ' 用于搜索的字符串,目前支持通配符的方式
        .Add "shtord", Array("nullable", "isInt", "positive")   ' 如果sheet order存在,则必须是正整数
        .Add "shtnm", Array("nullable")                         ' sheet name
        .Add "range", Array("nullable", "isPairRef")            ' 可以通过定义range限定范围
        .Add "comment", Array("nullable")                       ' 打开则只搜索comment,关闭则搜索正常内容.开关型参数不验证又不好,勉强验证一下
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

' https://docs.microsoft.com/en-us/office/vba/api/excel.range.find
const xlComments = -4144
const xlValues = -4163

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

    call afterCheck(xlFile, params, sheetParam)
end sub

sub afterCheck(byRef xlFile, params, sheetParam)
    ' 获取参数
    pat = params.item("pattern")
    rangeLimit = params.item("range")               ' 不要默认值
    commentInstead = params.exists("comment")       ' 开关型参数只看有没有,不看值是什么

    if params.exists("shtord") or params.exists("shtnm") then
        ' 若指定sheet,则限定sheet搜索
        set xlSheet = xlFile.worksheets(sheetParam)
        call searchOneSheet(xlSheet, pat, rangeLimit, commentInstead)
    else
        ' 若不指定,默认在所有sheet中搜索
        for each xlsheet in xlFile.worksheets
            call searchOneSheet(xlsheet, pat, rangeLimit, commentInstead)
        next
    end if
end sub

' 在一个sheet中搜索内容
' para: xlWorkbook, search pattern, range, is searching comment
sub searchOneSheet(byRef xlSheet, pat, rangeLimit, commentInstead)
    dim searchRange

    ' 若不指定搜索范围,默认使用UsedRange
    if len(rangeLimit) > 0 then
        set searchRange = xlSheet.Range(rangeLimit)
    else
        set searchRange = xlSheet.UsedRange
    end if

    set entries = findAll(searchRange, pat, commentInstead)
    if entries is nothing then
        wscript.echo("nothing found in " & xlsheet.name)
    else
        for each entry in entries
            if not commentInstead then
                wscript.echo(xlsheet.name & "!" & entry.Address & ": " & entry.Value)
            else
                wscript.echo(xlsheet.name & "!" & entry.Address & ": " & entry.Value & "(" & entry.Comment.Text & ")")
            end if
        next
    end if

    set searchRange = nothing
end sub

' 在所选范围内搜索所有结果
' para: search range, search pattern, is searching comment
' return: nothing if not found
' return: List{range1, range2, }
function findAll(byRef searchRange, pat, commentInstead)
    dim rng
    set res = createobject("system.collections.arraylist")

    lookInParam = xlValues
    if commentInstead then
        lookInParam = xlComments
    end if
    set rng = searchRange.Cells.Find(pat, , lookInParam)   ' vbs提供了很奇怪的API,一次只能搜索一个内容
    if not rng is nothing then
        firstAddress = rng.Address
        do
            res.Add(rng)
            ' 如果搜索范围只有一个格,会报错没有findnext方法,建议此时使用get而不是grep
            set rng = searchRange.FindNext(rng)   ' 接下来的搜索需要用FindNext(参数里还要有目前找到的结果)
        loop until rng.Address = firstAddress     ' 而且搜索到了结尾还不停止,会再次找到第一个
    else
        set res = nothing
    end if

    set findAll = res
end function

' 加载后的欢迎语句
wscript.echo("component: grep.vbs loaded")