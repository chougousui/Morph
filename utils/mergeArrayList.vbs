' 合并两个ArrayList(隶属于object)
' 考虑参数为空的情况
function mergeArrayList(arr1, arr2)
    set res = createObject("System.Collections.ArrayList")

    ' 参数为空,vartype 为 vbEmpty(0)
    if VarType(arr1) = vbObject then
        ' 参数为nothing也是object版本的空,也要过滤掉
        if not arr1 is nothing then
            for each item in arr1
                res.Add item
            next
        end if
    end if

    if VarType(arr2) = vbObject then
        if not arr2 is nothing then
            for each item in arr2
                res.Add item
            next
        end if
    end if

    set mergeArrayList = res
end function

'----------------------------------
if wscript.arguments.Count = 0 then
    set l1 = createObject("System.Collections.ArrayList")

    set l2 = createObject("System.Collections.ArrayList")
    l1.add Array("requiredAny", Array("testInt", "testString"))

    set l3 = mergeArrayList(l1, l2)
    wscript.echo(l3.Count)
    wscript.echo(l3.item(0)(0))
end if