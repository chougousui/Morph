' 为了让读dict行为,真正地成为只读
function onlyReadItem(byRef dict, key)
    if dict.exists(key) then
        ' 单独使用 dict.item()会导致不存在的key突然存在
        onlyReadItem = dict.item(key)
    end if
end function