' 执行具体的工作
' 1. 命令行参数的验证(从其他文件读取验证规则,处理函数)
' 2. 从其他文件读取配置的功能
' 3. 处理文件的打开关闭,错误处理
' component或许不是个好名字,worker?service?

class component
    ' 以列表,字典等保存配置
    private m_subCommand
    private m_mixins
    private m_mixinObjects
    private m_example
    private m_readonly
    private m_validateRules
    private m_crossValidateRules

    public property get example
        msg = "e.g.(bash):" & vbcrlf _
          & m_example.item("bash") & vbcrlf _
          & "e.g.(powershell):" & vbcrlf _
          & m_example.item("powershell")
        example = msg
    end property

    ' 初始化时include一些外部文件,同时初始化空的列表和字典
    private sub class_Initialize
        ' Called automatically when class is created
        Include(resolvePath("./utils/mergeDict.vbs"))
        Include(resolvePath("./utils/mergeArrayList.vbs"))
        Include(resolvePath("./partialComponent/staticParamValidation.vbs"))
        Include(resolvePath("./AppManager.vbs"))
        Include(resolvePath("./WorkbookManager.vbs"))

        set m_mixinObjects = createobject("System.Collections.ArrayList")
        set m_example = createObject("Scripting.Dictionary")
        with m_example
            .Add "bash", ""
            .Add "powershell", ""
        end with
        m_readonly = false
        set m_validateRules = createObject("Scripting.Dictionary")
        set m_crossValidateRules = createObject("System.Collections.ArrayList")
    end sub

    ' 销毁类时也做一些资源清理
    ' 此时若报错则清理不能完成
    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
        set m_mixinObjects = nothing
        set m_example = nothing
        set m_validateRules = nothing
        set m_crossValidateRules = nothing
    end sub

    ' 根据子命令找到散落在其他文件中的函数并加载
    public default function constructor(subCommand)
        m_subCommand = subCommand

        ' 依赖类成员 m_subCommand 的值
        call loadPartialComponent()

        ' 依赖类成员 m_mixins 的值
        call loadMixins()

        set constructor = me
    end function

    ' 将分布在其他文件中的partial类中的内容合并过来
    private sub loadPartialComponent()
        Include(resolvePath("./partialComponent/" & m_subCommand & ".vbs"))

        ' 先读mixins
        m_mixins = configs().item("mixins")

        ' 读取用例
        set m_example = mergeDict(m_example, configs().item("example"))

        ' 只读属性只来自于具体的subcommand对应的component
        m_readonly = configs().item("readonly")

        ' 验证规则
        set m_validateRules = mergeDict(m_validateRules, configs().item("validationRules"))
        set m_crossValidateRules = mergeArrayList(m_crossValidateRules, configs().item("crossValidationRules"))
    end sub

    ' 从mixin中合并一些配置(用例,验证规则,交叉验证规则)
    private sub loadMixins()
        ' 若组件中没有生命mixin则不再继续处理
        ' vbEmpty = 0, vbArray > 8192
        if VarType(m_mixins) < vbArray then
            exit sub
        end if

        for each mixinName in m_mixins
            Include(resolvePath("./mixins/" & mixinName & ".vbs"))
            set tempInstance = getMixinTempInstance()
            ' 保存instance,后面的layeredProcess还要再次用到mixin的处理
            m_mixinObjects.Add tempInstance

            set mixinConfigs = tempInstance.configs
            set m_example = mergeDict(m_example, mixinConfigs.item("example"))
            set m_validateRules = mergeDict(m_validateRules, mixinConfigs.item("validationRules"))
            set m_crossValidateRules = mergeArrayList(m_crossValidateRules, mixinConfigs.item("crossValidationRules"))
        next
    end sub

    ' 工作核心
    ' 为了读取和存储其他文件中的信息,其他部分已经写得太长太长
    public sub processFiles(files)
        ' 准备用于layeredProcess的参数,将只读的命令行参数,誊到一个dictionary中
        set params = createObject("Scripting.Dictionary")
        for each key in wscript.arguments.Named
            params.item(key) = wscript.arguments.Named.item(key)
        next

        ' 做静态参数校验(子命令相关,文件无关)
        call staticParamValidation(params, m_validateRules, m_crossValidateRules, example)

        ' 用对象封装App,借生命周期钩子,自动退出App
        set aManager = new AppManager

        for each file in files: do
            ' 用文件管理对象方便操作
            set wbManager = (new WorkbookManager)(file, m_readonly)

            ' 文件自己管理状态和输出内容
            wbManager.prompt()

            ' 打开文件
            call wbManager.openFileByApp(aManager.xlApp)

            ' 处理文件
            ' 即使处理失败,也想关闭文件
            on error resume next
            wscript.echo("---------------------------------------")
            call processOneFile(wbManager, params)
            wscript.echo("---------------------------------------")

            ' 关闭文件
            if err.number = 0 then
                call wbManager.closeFile()
            else
                call wbManager.resetFile()
            end if
            on error goto 0
        loop while false: next

        call aManager.quitApp()
    end sub

    private sub processOneFile(byRef wbManager, byRef params)
        ' 为了在出错时显示更多信息而继续
        ' 但遇到错误就退出处理当前文件的思想不变,在错误处理中直接退出
        on error resume next

        ' 依次执行mixin中的layered process
        for each mixinObj in m_mixinObjects
            set params = mixinObj.layeredProcess(wbManager.xlFile, params)
            if err.number <> 0 then
                wscript.echo("error occured in mixin " & mixinObj.name & ", #" _
                            & err.number & ", " & err.description)
                exit sub
            end if
        next

        ' comp中的layered process
        set params = layeredProcess(wbManager.xlFile, params)

        ' comp中的readwork
        call realWork(wbManager.xlFile, params)

        if err.number <> 0 then
            wscript.echo("error occured in component " & m_subCommand & ", #" _
                         & err.number & ", " & err.description)
            exit sub
        end if
    end sub

end class


' ------------------------------------------------