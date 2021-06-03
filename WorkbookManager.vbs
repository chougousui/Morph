' 主要用于将文件对象与readonly属性绑定在一起
' 清理资源是次要,在其他文件里也能做

class WorkbookManager
    private m_filename
    private m_xlFile
    private m_readonly

    public property get readonly
        readonly = m_readonly
    end property

    public property get xlFile
        set xlFile = m_xlFile
    end property
    public property let xlFile(Value)
        set m_xlFile = value
    end property

    private sub class_Initialize
        ' Called automatically when class is created
        m_readonly = false
        set m_xlFile = nothing
    end sub

    public default function constructor(pFilename, pReadonly)
        m_filename = pFilename
        m_readonly = pReadonly

        set constructor = me
    end function

    private sub class_Terminate
        ' Called automatically when all references to class instance are removed
        resetFile()
    end sub

    public sub openFileByApp(byRef xlApp)
        ' app竟然是字符串
        if not xlApp is nothing then
            set m_xlFile = xlApp.workbooks.open(m_filename, , m_readonly)
        end if
    end sub

    public function prompt()
        readonlyMark = "[edit]"
        if m_readonly then
            readonlyMark = "[readonly]"
        end if
        wscript.echo(vbcrlf & m_filename & " " & readonlyMark)
    end function

    public sub closeFile()
        if not m_xlFile is nothing then
            needSave = not m_readonly
            ' true表示保存内容,以只读模式打开时,会弹出页面提示
            m_xlFile.close(needSave)
            if needSave then
                wscript.echo("saved")
            end if
            set m_xlFile = nothing
        end if
    end sub

    public sub resetFile()
        if not m_xlFile is nothing then
            ' false表示不保存内容
            m_xlFile.close(false)
            set m_xlFile = nothing

            wscript.echo("closed without saving")
        end if
    end sub
end class