class rangeMixin
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
        m_name = "rangeMixin"

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
            .Add "bash", "/range:A1"
            .Add "powershell", "/range:A1"
        end with

        ' 可选,参数的验证规则
        ' Dict{{opt1, (rule1, rule2,)},{opt2, (rule1, rule2,)},}
        set validationRules = createObject("Scripting.Dictionary")
        with validationRules
            .Add "range", Array("required", "isSingleRef")          ' 必须有
        end with

        set m_configs = createObject("Scripting.Dictionary")
        with m_configs
            .Add "example", example
            .Add "validationRules", validationRules
        end with
    end sub

    ' 必须
    public function layeredProcess(byRef xlFile, byRef params)
        ' 针对range,无法做更多的与文件相关的验证,因此直接传递
        set layeredProcess = params
    end function

end class

function getMixinTempInstance()
    set getMixinTempInstance = new rangeMixin
end function

'----------------------------------