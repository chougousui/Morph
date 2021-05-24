' 验证函数集合
' 限于vbs无法include,将功能代码与测试代码放在一起

'--------------------------------
' 简谱的assert快捷键
sub assertToBeTrue(a)
    call assertToBe(a, true)
end sub

sub assertToBeFalse(a)
    call assertToBe(a, false)
end sub

sub assertToBe(a, b)
    if a <> b then
        wscript.echo("error, assert " & a & " to be " & b)
    else
        wscript.echo("pass, " & a & " is " & b)
    end if
end sub
'--------------------------------
' 测试功能列表
sub testAll()
    testRequired()
    testIsInt()
    testPositive()
    testConflict()
    testIsRef()
    testAtMostPairRef()
end sub
'--------------------------------
' 目标字符串不能为空
function required(value)
    required = len(value) > 0
end function

sub testRequired()
    v1 = required("")
    assertToBeFalse(v1)

    v2 = required("1")
    assertToBeTrue(v2)
end sub
'--------------------------------
' 如果为空则不再进行验证
function nullable(value)
    ' 函数内部仍检查是否为空并返回
    nullable = len(value) > 0
end function
'--------------------------------
' 必须是整数
function isInt(value)
    ' vbs逻辑与求值无法只在第一项为false时停下
    if IsNumeric(value) then
        if cstr(clng(value)) = value then
            isInt = true
            exit function
        end if
    end if
    isInt = false
end function

sub testIsInt()
    v3 = isInt("1")
    assertToBeTrue(v3)

    v4 = isInt("01")
    assertToBeFalse(v4)

    v5 = isInt("1.2")
    assertToBeFalse(v5)

    v6 = isInt("a")
    assertToBeFalse(v6)

    v7 = isInt("1,2,3")
    assertToBeFalse(v7)
end sub
'--------------------------------
' 数字大于0
function positive(value)
    if IsNumeric(value) then
        if cdbl(value) > 0 then
            positive = true
            exit function
        end if
    end if
    positive = false
end function

sub testPositive()
    v8 = positive("a")
    assertToBeFalse(v8)

    v9 = positive("0")
    assertToBeFalse(v9)

    v10 = positive("0.1")
    assertToBeTrue(v10)

    v11 = positive("1")
    assertToBeTrue(v11)
end sub
'--------------------------------
' 交叉验证: 两参数冲突,只能指定一个
function conflict(params)
    ' params 是一个二维数组[[参数名, 参数是否存在, 参数值]]
    ' 目前 conflict 要求有且只有两个参数,否则返回false
    if params.Count <> 2 then
        conflict = false
        exit function
    end if

    conflict = not (params.item(0)(1) and params.item(1)(1))
end function

sub testConflict()
    set ctx = CreateObject("System.Collections.ArrayList")
    with ctx
        .Add Array("shtord", true, "1")
        .Add Array("shtnm", true, "string")
    end with

    v12 = conflict(ctx)
    assertToBeFalse(v12)

    set ctx = CreateObject("System.Collections.ArrayList")
    with ctx
        .Add Array("shtord", true, "2")
        .Add Array("shtnm", false, "")
    end with

    v13 = conflict(ctx)
    assertToBeTrue(v13)

    set ctx = CreateObject("System.Collections.ArrayList")
    with ctx
        .Add Array("shtord", false, "")
        .Add Array("shtnm", false, "")
    end with

    v14 = conflict(ctx)
    assertToBeTrue(v14)
end sub
'--------------------------------
' 判断range表达式是否合法表示一个格
function isSingleRef(value)
    set oRe = new RegExp
    oRe.Pattern = "^[A-Z]{1,2}[0-9]{1,5}$"
    isSingleRef = oRe.Test(value)
end function

sub testIsRef()
    v15 = isSingleRef("A1")
    assertToBeTrue(v15)

    v16 = isSingleRef("1A")
    assertToBeFalse(v16)

    v17 = isSingleRef("123")
    assertToBeFalse(v17)

    ' https://www.vishalon.net/blog/excel-column-letter-to-number-quick-reference#columnzazz
    v18 = isSingleRef("IV65536")
    assertToBeTrue(v18)

    ' v18 = isRef("XFD1048576")
    ' assertToBeTrue(v18)
    v19 = isSingleRef("A1:B2")
    assertToBeFalse(v19)
end sub
'--------------------------------
' 判断range表达式是否是两个合法ref
function isPairRef(value)
    frags = Split(value, ":")

    ' Ubound得到数组的最大合法下标,+1后表示数组长度
    if Ubound(frags) + 1 <> 2 then
        isPairRef = false
        exit function
    end if

    res = true
    for each frag in frags
        res = res and isSingleRef(frag)
    next

    res = res and (frags(0) <> frags(1))

    isPairRef = res
end function

sub testAtMostPairRef()
    v20 = isPairRef("A1")
    assertToBeFalse(v20)

    v21 = isPairRef("A1:B2")
    assertToBeTrue(v21)

    v22 = isPairRef("123AB")
    assertToBeFalse(v22)

    v23 = isPairRef(":")
    assertToBeFalse(v23)

    v24 = isPairRef(":B1")
    assertToBeFalse(v24)

    v25 = isPairRef("A1:")
    assertToBeFalse(v25)

    v26 = isPairRef("AA:B1")
    assertToBeFalse(v26)

    v27 = isPairRef("A1:B1:")
    assertToBeFalse(v27)
end sub

' -------------------------------------------------------------------------
' 如果被include,将共用morph.vbs的参数列表,即 0: <subCommand>, 1: <wildcard> 2~: [options]
' 如果没有参数,说明是直接测试,需要运行测试函数
if wscript.arguments.Unnamed.Count = 0 then
    testAll()
end if
' 测试时使用   cscript ./validators.vbs | tail -n+4 | cat -n