' 合并两个Dictionary
' 考虑参数为空的情况
' 考虑key重复情况
function mergeDict(dict1, dict2)
    set res = createObject("Scripting.Dictionary")

    ' 参数为空,vartype 为 vbEmpty(0)
    if VarType(dict1) = vbObject then
        ' 参数为nothing也是object版本的空,也要过滤掉
        if not dict1 is nothing then
            for each key in dict1
                res.Add key, dict1.item(key)
            next
        end if
    end if

    if VarType(dict2) = vbObject then
        if not dict2 is nothing then
            for each key in dict2
                ' 考虑key重复
                if res.exists(key) then
                    ' 若类型不同,无法合并
                    if VarType(res.item(key)) <> VarType(dict2.item(key)) then
                        call err.raise(13, "", "cannot merge different type into same key")
                    end if

                    ' 不是object,就能不使用set关键字
                    if vartype(dict2.item(key)) <> vbObject then
                        firstValue = res.item(key)
                        secondValue = dict2.item(key)

                        ' 若均为 string,则按照规则合并
                        if VarType(secondValue) = vbString then
                            mergedValue = mergeString(firstValue, secondValue)
                        end if

                        ' 若均为 Array, 也按规则合并
                        ' TODO 目前暂时无法处理更深层次的类型,Array(string, Array) 与 Array(Array) 无法区分
                        if vartype(dict2.item(key)) > vbArray then
                            ' 如果value为array,则合并两个array后赋值
                            mergedValue = mergeArray(firstValue, secondValue)
                        end if
                    else
                        ' 若是object,这里的语句需要带上set关键字
                        set firstValue = res.item(key)
                        set secondValue = dict2.item(key)

                        ' 但目前计划先丢弃冲突的object
                        ' set mergedValue = mergeObject(firstValue, secondValue)
                        set mergedValue = nothing
                    end if

                    res.item(key) = mergedValue
                else
                    ' key不重复则直接添加
                    res.Add key, dict2.item(key)
                end if
            next
        end if
    end if

    set mergeDict = res
end function

function mergeString(str1, str2)
    res = ""
    if VarType(str1) = vbString then
        res = str1
    end if

    if VarType(str2) = vbString then
        if len(res) > 0 then
            res = res & " " & str2
        else
            res = str2
        end if
    end if

    mergeString = res
end function

function mergeArray(arr1, arr2)
    ' vbs 的 array, length = Ubound(arr) + 1
    ' vbs 的 array(3) 可以放 4 个元素
    ' dim 还只能以常量表示长度, redim 才能用变量表示
    redim res(Ubound(arr1) + Ubound(arr2) + 1)
    for i = 0 to Ubound(arr1)
        res(i) = arr1(i)
    next
    for i = 0 to Ubound(arr2)
        res(Ubound(arr1)+1+i) = arr2(i)
    next

    mergeArray = res
end function

'--------------------------------
sub printArray(arr)
    wscript.echo("----------")
    wscript.echo("length " & Ubound(arr) + 1)
    wscript.echo("----------")
    for each item in arr
        wscript.echo(item)
    next
    wscript.echo("----------")
end sub

sub printDict(dict)
    wscript.echo("----------")
    for each key in dict
        wscript.echo("key: " & key)
        wscript.echo("value: " & dict.item(key))
    next
    wscript.echo("----------")
end sub

if wscript.arguments.Count = 0 then
    set dict1 = createobject("scripting.Dictionary")
    with dict1
        .Add "bash", ""
        .Add "powershell", ""
    end with

    set dict2 = createObject("Scripting.Dictionary")
    with dict2
        .Add "bash", "morph eg ./assets/test*/*.xlsx /testInt:1 /testString:A"
        .Add "powershell", "morph eg .\assets\test*\*.xlsx /testInt:1 /testString:A"
    end with

    set dict3 = mergeDict(dict1, dict2)
    printDict(dict3)
end if