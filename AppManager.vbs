' 借助class_Terminate自动关闭文件和APP
class AppManager
    private m_xlApp
    public property get xlApp
        set xlApp = m_xlApp
    end property

    private sub class_Initialize
        ' Called automatically when class is created
        ' 不敢相信xlApp是字符串类型的,竟然还能操作...
        set m_xlApp = createobject("Excel.Application")
    end sub

    private sub class_Terminate
        ' Called automatically when all references to class xlApp are removed
        call quitApp()
    end sub

    public sub quitApp
        if not m_xlApp is nothing then
            ' 关闭app前先检查打开的文件,并不保存关闭
            counter = m_xlApp.workbooks.Count
            if counter > 0 then
                wscript.echo(counter & "files not closed")
            end if

            for each workbook in m_xlApp.workbooks
                ' 关闭workbook后无法再读取name
                filename = workbook.name
                ' false表示不保存
                workbook.Close(false)
                wscript.echo("file " & filename & " closed without saving")
                set workbook = nothing
            next

            ' 然后再退出App
            m_xlApp.Quit
            wscript.echo("app quitted")
            set m_xlApp = nothing
        end if
    end sub

end class