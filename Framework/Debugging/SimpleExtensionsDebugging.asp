<%
'''
 ' SimpleExtensionsDebugging.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.14
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsDebugging

    Private dbg_Enabled
    Private dbg_Show
    Private dbg_RequestDate
    Private dbg_RequestTimer
    Private dbg_FinishTimer
    Private dbg_Data
    Private dbg_DB_Data
    Private dbg_Show_default
    Private divSets(2)

    'Construktor => set the default values
    Private Sub Class_Initialize()
        dbg_RequestDate = Now()
        dbg_RequestTimer = Timer()
        Set dbg_Data = Server.CreateObject("Scripting.Dictionary")
        divSets(0) = "<tr><td style=""cursor:pointer;"" onclick=""javascript:if (document.getElementById('data#sectname#').style.display=='none'){document.getElementById('data#sectname#').style.display='block';}else{document.getElementById('data#sectname#').style.display='none';}""><div id=""sect#sectname#"" style=""font-weight:bold;cursor:pointer;background:#7EA5D7;color:white;padding-left:4px;padding-right:4px;padding-bottom:2px;"">|#title#|  <div id=data#sectname# style=""cursor:text;display:none;background:#FFF;padding-left:8px;"" onclick=""window.event.cancelBubble = true;"">|#data#|  </div>|</div>|"
        divSets(1) = "<tr><td><div id=""sect#sectname#"" style=""font-weight:bold;cursor:pointer;background:#7EA5D7;color:white;padding-left:4px;padding-right:4px;padding-bottom:2px;"" onclick=""javascript:if (document.getElementById('data#sectname#').style.display=='none'){document.getElementById('data#sectname#').style.display='block';}else{document.getElementById('data#sectname#').style.display='none';}"">|#title#|  <div id=data#sectname# style=""cursor:text;display:block;background:#FFF;padding-left:8px;"" onclick=""window.event.cancelBubble = true;"">|#data#|  </div>|</div>|"
        divSets(2) = "<tr><td><div id=sect#sectname# style=""background:#7EA5D7;color:lightsteelblue;padding-left:4px;padding-right:4px;padding-bottom:2px;"">|#title#|  <div id=data#sectname# style=""display:none;background:lightsteelblue;padding-left:8px"">|#data#|  </div>|</div>|"
        dbg_Show_default = "0,0,0,0,0,0,0,0,0,0,0"
    End Sub

    '''
     ' 启用调试
     ''
    Public Function enabled()
        dbg_Enabled = True
    End Function

    '''
     ' 关闭调试
     ''
    Public Function disable()
        dbg_Enabled = False
    End Function

    '''
     ' 设置面板状态
     ' 
     ' @param string panelsStatusString <面板状态字符串，
     ' 11个面板按顺序以逗号分割
     ' 0：隐藏
     ' 1：显示
     ' 例："1,0,0,0,0,0,0,0,0,0,0">
     ''
    Public Function setPanelsStatus(ByVal panelsStatusString)
        dbg_Show = panelsStatusString
    End Function

    '******************************************************************************************************************
    ''@SDESCRIPTION: Adds a variable to the debug-informations.
    ''@PARAM:   - variableName [string]: Description of the variable
    ''@PARAM:   - variable [variable]: The variable itself
    '******************************************************************************************************************
    Public Sub setVariable(ByVal variableName, ByVal variable)
        If dbg_Enabled Then
            If Err.Number > 0 Then
                Call dbg_Data.Add(validLabel(variableName), "!!! Error: " & Err.Number & " " &  Err.Description)
                Err.Clear
            Else
                Dim uniqueID
                uniqueID = validLabel(variableName)
                Call dbg_Data.Add(uniqueID, variable)
            End If
        End If
    End Sub

    '******************************************************************************************************************
    ''@SDESCRIPTION: Draws the Debug-panel
    '******************************************************************************************************************
    Public Sub draw()
        If dbg_Enabled Then
            dbg_FinishTimer = Timer()

            Dim divSet, x
            divSet = Split(dbg_Show_default,",")
            dbg_Show = Split(dbg_Show,",")

            For x = 0 To UBound(dbg_Show)
                divSet(x) = dbg_Show(x)
            Next

            Response.Write("<br><table width=100% cellspacing=0 border=0 style=""font-family:arial;font-size:9px;font-weight:normal;""><tr><td><div style=""background:#005A9E;color:white;padding:4px;font-size:16px;font-weight:bold;"">调试输出:</div>")
            Call printSummaryInfo(divSet(0))
            Call printCollection("变量", dbg_Data,divSet(1), "")
            Call printCollection("QueryString 集合", Request.QueryString(), divSet(2), "")
            Call printCollection("Form 集合", Request.Form(), divSet(3), "")
            Call printCookiesInfo(divSet(4))
            Call printCollection("Session", Session.Contents(), divSet(5),AddRow(AddRow(AddRow("", "Locale ID", Session.LCID & " (&H" & Hex(Session.LCID) & ")"), "Code Page",Session.CodePage), "Session ID", Session.SessionID))
            Call printCollection("Application 对象", Application.Contents(), divSet(6), "")
            Call printCollection("服务器变量", Request.ServerVariables(), divSet(7), AddRow("","Timeout",Server.ScriptTimeout))
            Call printDatabaseInfo(divSet(8))
            Call printCollection("Session StaticObjects 集合", Session.StaticObjects(), divSet(9), "")
            Call printCollection("Application StaticObjects 集合", Application.StaticObjects(), divSet(10), "")
            Response.Write("</table>")
        End If
    End Sub

    '******************************************************************************************************************
    ''@SDESCRIPTION: Adds the Database-connection object to the debug-instance. To display Database-information
    ''@PARAM:   - oSQLDB [object]: connection-object
    '******************************************************************************************************************
    Public Sub setDatabaseInfo(ByVal oSQLDB)
        dbg_DB_Data = addRow(dbg_DB_Data, "ADO Ver", oSQLDB.Version)
        dbg_DB_Data = addRow(dbg_DB_Data, "OLEDB Ver", oSQLDB.Properties("OLE DB Version"))
        dbg_DB_Data = addRow(dbg_DB_Data, "DBMS", oSQLDB.Properties("DBMS Name") & " Ver: " & oSQLDB.Properties("DBMS Version"))
        dbg_DB_Data = addRow(dbg_DB_Data, "Provider", oSQLDB.Properties("Provider Name") & " Ver: " & oSQLDB.Properties("Provider Version"))
    End Sub

    '******************************************************************************************************************
    '* ValidLabel
    '******************************************************************************************************************
    Private Function validLabel(ByVal label)
        Dim i, lbl
        i = 0
        lbl = label
        Do
            If Not dbg_Data.Exists(lbl) Then Exit Do
            i = i + 1
            lbl = label & "(" & i & ")"
        Loop Until i = i

        validLabel = lbl
    End Function

    '******************************************************************************************************************
    '* PrintCookiesInfo
    '******************************************************************************************************************
    Private Sub printCookiesInfo(ByVal divSetNo)
        Dim tbl, cookie, key, tmp
        For Each cookie In Request.Cookies
            If Not Request.Cookies(cookie).HasKeys Then
                tbl = AddRow(tbl, cookie, Request.Cookies(cookie))
            Else
                For Each key In Request.Cookies(cookie)
                    tbl = AddRow(tbl, cookie & "(" & key & ")", Request.Cookies(cookie)(key))
                Next
            End If
        Next

        tbl = makeTable(tbl)
        If Request.Cookies.Count <= 0 Then divSetNo = 2
        tmp = Replace(Replace(Replace(divSets(divSetNo),"#sectname#","COOKIES"),"#title#","Cookies"),"#data#",tbl)
        Response.Write(Replace(tmp,"|", vbcrlf))
    End Sub

    '******************************************************************************************************************
    '* PrintSummaryInfo
    '******************************************************************************************************************
    Private Sub printSummaryInfo(ByVal divSetNo)
        Dim tmp, tbl
        tbl = addRow(tbl, "请求时间", dbg_RequestDate)
        tbl = addRow(tbl, "耗时", FormatNumber((dbg_FinishTimer - dbg_RequestTimer), 10) & " 秒")
        tbl = addRow(tbl, "请求类型", Request.ServerVariables("REQUEST_METHOD"))
        tbl = addRow(tbl, "服务器状态", Response.Status)
        tbl = addRow(tbl, "脚本引擎", ScriptEngine & " " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion)
        tbl = makeTable(tbl)
        tmp = Replace(Replace(Replace(divSets(divSetNo), "#sectname#", "SUMMARY"), "#title#", "摘要信息"), "#data#", tbl)
        Response.Write(Replace(tmp,"|", vbCrlf))
    End Sub

    '******************************************************************************************************************
    '* printDatabaseInfo
    '******************************************************************************************************************
    Private Sub printDatabaseInfo(ByVal divSetNo)
        Dim tbl
        tbl = makeTable(dbg_DB_Data)
        tbl = replace(replace(replace(divSets(divSetNo),"#sectname#", "DATABASE"), "#title#", "DATABASE INFO"), "#data#", tbl)
        Response.Write replace(tbl, "|", vbCrlf)
    End Sub

    '******************************************************************************************************************
    '* printCollection
    '******************************************************************************************************************
    Private Sub printCollection(Byval name, ByVal collection, ByVal divSetNo, ByVal extraInfo)
        Dim vItem, tbl, Temp
        For Each vItem In collection
            If IsObject(collection(vItem)) And name <> "服务器变量" And name <> "QueryString 集合" And name <> "Form 集合" Then
                tbl = addRow(tbl, vItem, "{Object}")
            ElseIf IsNull(collection(vItem)) Then
                tbl = addRow(tbl, vItem, "{Null}")
            ElseIf IsArray(collection(vItem)) Then
                tbl = addRow(tbl, vItem, "{Array}")
            Else
                If (name = "服务器变量" And vItem <> "ALL_HTTP" And vItem <> "ALL_RAW") Or name <> "服务器变量" Then
                    If collection(vItem) <> "" Then
                        tbl = addRow(tbl, vItem, Server.HTMLEncode(collection(vItem))) ' & " {" & TypeName(collection(vItem)) & "}")
                    Else
                        tbl = addRow(tbl, vItem, "...")
                    End If
                End If
            End If
        Next
        If extraInfo <> "" Then tbl = tbl & "<tr><td colspan=2><hr></tr>" & extraInfo
        tbl = makeTable(tbl)
        If collection.count <= 0 Then divSetNo = 2
        tbl = Replace(Replace(divSets(divSetNo), "#title#", Name), "#data#", tbl)
        tbl = Replace(tbl, "#sectname#", Replace(Name, " ", ""))
        Response.Write Replace(tbl, "|", vbCrlf)
    End Sub

    '******************************************************************************************************************
    '* addRow
    '******************************************************************************************************************
    Private Function addRow(ByVal t, ByVal var, ByVal val)
        t = t & "|<tr valign=""top"" style=""color:#000;"">|<td>|" & var & "|<td>= " & val & "|</tr>"
        addRow = t
    End Function

    '******************************************************************************************************************
    '* makeTable
    '******************************************************************************************************************
    Private Function makeTable(ByVal tdata)
        tdata = "|<table style=""border:0;font-size:13px;font-weight:normal;"">" + tdata + "</table>|"
        makeTable = tdata
    End Function

    'Destructor
    Private Sub Class_Terminate()
        Set dbg_Data = Nothing
    End Sub

End Class
%>