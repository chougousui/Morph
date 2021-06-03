' 搜索文本
' 1. 在excel中搜索想要的文本
' 2. 目前支持通配符搜索
' 3. 能够指定sheet和range缩小搜索范围
' 4. 使用开关决定是否转去搜索comment


function configs()
    ' 必要,组件的用例
    ' Dict{"bash", "powershell"}
    set example = createObject("Scripting.Dictionary")
    with example
        .Add "bash", "morph grep ./assets/test*/*.txt /pattern:""a?c"" [/range:""A1:F10""] [/comment:]"
        .Add "powershell", "morph grep .\assets\test*\*.txt /pattern:""a?c"" [/range:""A1:F10""] [/comment]"
    end with

    ' 必要,组件的操作权限
    ' true/false
    readonly = true

    ' 可选,混入的其他可独立处理过程
    ' Array(mixin1, mixin2,)
    mixins = Array("sheetMixin")

    ' 可选,参数的验证规则
    ' Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
    set validationRules = createObject("Scripting.Dictionary")
    with validationRules
        .Add "pattern", Array("required")                       ' 用于搜索的字符串,目前支持通配符的方式
        .Add "range", Array("nullable", "isPairRef")            ' 可以通过定义range限定范围
        .Add "comment", Array("nullable")                       ' 打开则只搜索comment,关闭则搜索正常内容.开关型参数不验证又不好,勉强验证一下
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
function layeredProcess(byRef xlFile, params)
    ' 两个默认值: 所有sheet, usedrange, 都无法在此处设置
    ' 所以此处不做任何事
    set layeredProcess = params
end function

' 必要,根据参数要求,执行操作
sub realWork(byRef xlFile, params)
    ' 取参数
    pat = params.item("pattern")
    sheetParam = onlyReadItem(params, "sheetParam")
    rangeLimit = onlyReadItem(params, "range")      ' 不要默认值
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

    ' 自行实现一个findall函数
    set entries = findAll(searchRange, pat, commentInstead)

    ' 显示查询结果
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

' https://docs.microsoft.com/en-us/office/vba/api/excel.range.find
const xlComments = -4144
const xlValues = -4163
' 在所选范围内搜索所有结果
' para: search range, search pattern, is searching comment
' return: nothing if not found
' return: List{range1, range2, }
function findAll(byRef searchRange, pat, commentInstead)
    dim rng
    set res = createobject("system.collections.arraylist")

    ' 默认搜索"值"
    lookInParam = xlValues
    if commentInstead then
        ' 有选项有搜索"注释"
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


' 为了让读dict行为,真正地成为只读
function onlyReadItem(byRef dict, key)
    if dict.exists(key) then
        ' 单独使用 dict.item()会导致不存在的key突然存在
        onlyReadItem = dict.item(key)
    end if
end function