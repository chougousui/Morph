' 包装组件中的内容,尤其是将组件中的规则,转化为一系列的处理函数
' 同时做一些格式化输出
' 为了解耦,内部不再包含错误处理代码

' 格式化组件的用例
' return: string
function wrappedExample()
    msg = "e.g.(bash):" & vbcrlf _
          & example().item("bash") & vbcrlf _
          & "e.g.(powershell):" & vbcrlf _
          & example().item("powershell")
    wrappedExample = msg
end function

' 根据可读性显示文件名的回显
' return: string
function wrappedFilePrompt(filename)
    readonlyMark = "[edit]"
    isReadOnly = readonly()
    if isReadOnly then
        readonlyMark = "[readonly]"
    end if
    wscript.echo(vbcrlf & filename & " " & readonlyMark)
end function

' 验证组件使用的参数
sub wrappedValidation(componentName, params)
    set validationResults = createobject("system.collections.arraylist")
    validated = true

    ' 单项验证
    set validationRulesF = myGetRef("validationRules", false)
    if not validationRulesF is nothing then
        set allRules = validationRulesF()
        call singleValidation(params, allRules, validated, validationResults)
    end if

    ' cross 验证
    set crossValidationRulesF = myGetRef("crossValidationRules", false)
    if not crossValidationRulesF is nothing then
        set allCrossRules = crossValidationRulesF()
        call crossValidation(params, allCrossRules, validated, validationResults)
    end if

    ' 只要有help就展示用法
    if params.exists("help") then
        validated = false
    end if

    if not validated then
        msg = "invalid arguments for component [" & componentName & "]" & vbcrlf _
              & vbcrlf _
              & "validation results:" & vbcrlf _
              & "---------------" & vbcrlf

        for each result in validationResults
            msg = msg & result & vbcrlf
        next

        msg = msg & "---------------" & vbcrlf _
                  & vbcrlf _
                  & wrappedExample()

        call err.raise(5, "", msg)
    end if
end sub

' 单项验证
sub singleValidation(params, allRules, byRef validated, byRef validationResults)
    for each key in allRules
        ' 获取一个字段对应的规则列表
        rules = allRules.item(key)
        ' 获取要判断的值
        value = params.item(key)

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
                validated = false
                msg = "[" & key & "]: rule [" & rule & "] not met, value: [" & value & "]"
                validationResults.add(msg)
            end if
        next
    next
end sub

' 交叉验证
sub crossValidation(params, allCrossRules, byRef validated, byRef validationResults)
    for each rule in allCrossRules
        ' 准备验证函数
        set f = myGetRef(rule, true)
        if f is nothing then
            call err.raise(35, "", "invalid cross validator " & rule)
        end if

        ' 准备用于验证的参数:一个只包含被验参数的(参数名,存在性,值)的二维数组
        ' 加强安全性的同时方便单独测试
        set ctx = CreateObject("System.Collections.ArrayList")
        optList = allCrossRules.item(rule)
        for each opt in optList
            ctx.add Array(opt, params.exists(opt), params.item(opt))
        next

        res = f(ctx)

        if not res then
            validated = false
            msg = "cross validation rule [" & rule & "] not met" & vbcrlf
            for each entry in ctx:
                msg = msg & "opt: " & entry(0) & ", exists: " & entry(1) & ", value: " & entry(2) & vbcrlf
            next
            validationResults.add(msg)
        end if
    next
end sub

' 组件的主要函数
sub wrappedExtension(xlFile, params)
    wscript.echo("---------------")
    call extension(xlFile, params)
    wscript.echo("---------------")
end sub

' 文件是否保存
' return: boolean
function wrappedNeedSave()
    wrappedNeedSave = not readonly()
end function