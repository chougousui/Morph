' 用于获取命令行参数的默认值
' para: value, defaultValue
' return: value if not empty string
' return: defaultValue
function orDefault(value, defaultValue)
    if len(value) > 0 then
        orDefault = value
        exit function
    end if
    orDefault = defaultValue
end function
