class sheetMixin
    private m_name
    private m_configs

    public property get name
        name = m_name
    end property

    public property get configs
        set configs = m_configs
    end property

    private sub class_Initialize
        ' 设置mixin的显示名,之后处理报错时,能够帮助定位问题
        m_name = "sheetMixin"

        call generateConfigs()
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
        set m_configs = nothing
    end sub

    private sub generateConfigs()
        ' 必要,组件的用例
        ' Dict{"bash", "powershell"}
        set example = createObject("Scripting.Dictionary")
        with example
            .Add "bash", "[/shtord:1 | /shtnm:""Sheet1""]"
            .Add "powershell", "[/shtord:1 | /shtnm:""Sheet1""]"
        end with

        ' 可选,参数的验证规则
        ' Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
        set validationRules = createObject("Scripting.Dictionary")
        with validationRules
            .Add "shtord", Array("nullable", "isInt", "positive")       ' 如果有,必须是正整数
            .Add "shtnm", Array("nullable")                             ' 可以没有
        end with

        ' 可选,参数的交叉验证规则
        ' ArrayList[(rule1, (opt1, opt2)), (rule2, (opt1, opt2, opt3,))]
        set crossValidationRules = CreateObject("System.Collections.ArrayList")
        with crossValidationRules
            .Add Array("conflict", Array("shtord", "shtnm"))     ' sheet order和sheet name不能同时存在
        end with

        set m_configs = createObject("Scripting.Dictionary")
        with m_configs
            .Add "example", example
            .Add "validationRules", validationRules
            .Add "crossValidationRules", crossValidationRules
        end with
    end sub

    ' 必须
    ' 此处获取sheet有关参数并做动态检查
    public function layeredProcess(byRef xlFile, byRef params)
        ' 当且仅当命令行指定了sheet相关参数
        ' 获取参数并进行文件内的校验
        if params.exists("shtord") or params.exists("shtnm") then
            sheetParam = getSheetParam(params)

            valid = validateSheetParam(xlFile, sheetParam)
            if not valid then
                call err.raise(5, "", "sheet params not valid in this file")
            end if

            params.item("sheetParam") = sheetParam
        end if

        set layeredProcess = params
    end function

    ' 从params中获取关于sheet的参数
    ' return: vbInteger or vbString or vbEmpty
    private function getSheetParam(params)
        if params.exists("shtord") then
            ' 需要转换成数字
            shtord = clng(params.item("shtord"))
            getSheetParam = shtord
        end if

        if params.exists("shtnm") then
            shtnm = params.item("shtnm")
            getSheetParam = shtnm
        end if
    end function

    ' 验证取到的sheet有关参数是否对于该文件合法
    ' return: true/false
    private function validateSheetParam(byRef xlFile, sheetParam)
        ' https://docs.microsoft.com/en-us/previous-versions/visualstudio/aa263402(v=vs.60)?redirectedfrom=MSDN
        ' const vbEmpty = 0
        ' const vbInteger = 2
        ' const vbLong = 3
        ' const vbString = 8

        ' 若为index,则验证是否超出范围,注意and操作符不会停止,不能写成一行if...
        if VarType(sheetParam) = vbInteger or VarType(sheetParam) = vbLong then
            if sheetParam > xlFile.worksheets.Count then
                wscript.echo("sheet order " & sheetParam & " not exists")
                validateSheetParam = false
                exit function
            end if
        end if

        ' 若为sheet名,则验证该sheet名是否存在
        if VarType(sheetParam) = vbString then
            ' 目前官方没有提供worksheets.exists()方法,API很不优雅,只能手动实现
            if not sheetNameExists(xlFile, sheetParam) then
                wscript.echo("sheet name " & sheetParam & " not exists")
                validateSheetParam = false
                exit function
            end if
        end if

        ' 其他情况,包括命令行没有指定sheet相关参数,都返回true
        validateSheetParam = true
    end function

    ' 判断sheet名是否存在
    ' return: true/false
    ' TODO 遍历法,试错法,isRef法
    private function sheetNameExists(byRef xlFile, sheetName)
        for each sheet in xlFile.worksheets
            if sheet.name = sheetName then
                sheetNameExists = true
                exit function
            end if
        next
        sheetNameExists = false
    end function

end class

function getMixinTempInstance()
    set getMixinTempInstance = new sheetMixin
end function

'----------------------------------