' 逻辑上作为component类的一部分,或称此处为partial class component
' 存放本该放在component类中的一些验证参数的具体实现
' 实质上由于vbs限制
' 1. 不能使用 comp.xxx()
' 2. 方法中不能直接使用类的成员,需要byRef传入

' 验证组件使用的参数
sub staticParamValidation(byRef params, byRef m_validateRules, byRef m_crossValidateRules, example)
    ' 初始化环境
    set validationResults = createobject("system.collections.arraylist")
    valid = true

    ' 准备验证函数
    ' 该相对路径竟然是相对于运行命令的路径
    Include(resolvePath("./validators.vbs"))

    ' 单项验证
    call singleValidation(params, m_validateRules, valid, validationResults)

    ' cross 验证
    call crossValidation(params, m_crossValidateRules, valid, validationResults)

    ' 只要有help就展示用法
    if params.exists("help") then
        valid = false
    end if

    ' 验证不通过时显示信息
    if not valid then
        msg = "invalid arguments for component [" & m_subCommand & "]" & vbcrlf _
              & vbcrlf _
              & "validation results:" & vbcrlf _
              & "---------------" & vbcrlf

        for each result in validationResults
            msg = msg & result & vbcrlf
        next

        msg = msg & "---------------" & vbcrlf _
                  & vbcrlf _
                  & example

        call err.raise(5, "", msg)
    end if
end sub

' 单项验证
sub singleValidation(params, allRules, byRef valid, byRef validationResults)
    for each key in allRules
        ' 获取一个字段对应的规则列表
        rules = allRules.item(key)
        ' 获取要判断的值
        value = onlyReadItem(params, key)

        for each rule in rules
            ' 取一个规则
            set f = myGetRef(rule, true)
            if f is nothing then
                call err.raise(35, "", "invalid validator " & rule)
            end if

            ' TODO if rule = "plurable", how to read next rule? queue?
            ' or solid functions like plurableIntPositive, plurableSingleRef

            ' 判断规则结果
            res = f(value)

            ' 如果规则nullable不满足则进行下一个选项的验证
            if rule = "nullable" and (not res) then
                exit for
            end if

            ' 如果出错则记录错误
            if not res then
                valid = false
                msg = "[" & key & "]: rule [" & rule & "] not met, value: [" & value & "]"
                validationResults.add(msg)
            end if
        next
    next
end sub

' 交叉验证
sub crossValidation(params, allCrossRules, byRef valid, byRef validationResults)
    for each line in allCrossRules
        rule = line(0)
        optList = line(1)

        ' 准备验证函数
        set f = myGetRef(rule, true)
        if f is nothing then
            call err.raise(35, "", "invalid cross validator " & rule)
        end if

        ' 准备用于验证的参数:一个只包含被验参数的(参数名,存在性,值)的二维数组
        ' 加强安全性的同时方便单独测试
        set ctx = CreateObject("System.Collections.ArrayList")
        for each opt in optList
            ctx.add Array(opt, params.exists(opt), onlyReadItem(params, key))
        next

        res = f(ctx)

        if not res then
            valid = false
            msg = "cross validation rule [" & rule & "] not met" & vbcrlf
            for each entry in ctx:
                msg = msg & "opt: " & entry(0) & ", exists: " & entry(1) & ", value: " & entry(2) & vbcrlf
            next
            validationResults.add(msg)
        end if
    next
end sub

'---------------------------------
' 为了让读dict行为,真正地成为只读
function onlyReadItem(byRef dict, key)
    if dict.exists(key) then
        ' 单独使用 dict.item()会导致不存在的key突然存在
        onlyReadItem = dict.item(key)
    end if
end function