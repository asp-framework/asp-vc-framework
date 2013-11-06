<%
'''
 ' SimpleExtensionsDBAccess.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.6
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsDBAccess

    ' @var dictionary <连接状态值>
    Private objectStateEnum

'###########################'
'###########################'

    Private Sub Class_Initialize
        ' 初始化参数
        initObjectStateEnum()
    End Sub

    '''
     ' 初始化连接状态值
     ''
    Private Sub initObjectStateEnum()
        Set objectStateEnum = Server.CreateObject("Scripting.Dictionary")
        Call objectStateEnum.Add("adStateClosed", 0)
        Call objectStateEnum.Add("adStateOpen", 1)
        Call objectStateEnum.Add("adStateConnecting", 2)
        Call objectStateEnum.Add("adStateExecuting", 4)
        Call objectStateEnum.Add("adStateFetching", 8)
    End Sub

'###########################'
'###########################'

    '''
     ' 打开数据库
     ''
    Public Function open()
        If SE.module("DB").getDBConnection.State <> objectStateEnum.Item("adStateClosed") Then _
            Exit Function

        open = 0
        Call SE.module("DB").getDBConnection.Open( _
            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & SE.module("DB").getDBSource & ";" & _
            "User Id=" & SE.module("DB").getDBUserName & ";" & _
            "Password=" & SE.module("DB").getDBPassword & ";" _
        )
        If SE.module("DB").getDBConnection.State = objectStateEnum.Item("adStateOpen") Then _
            open = 1
    End Function

    '''
     ' 关闭数据库
     ''
    Public Function close()
        If SE.module("DB").getDBConnection.State = objectStateEnum.Item("adStateClosed") Then _
            Exit Function

        close = 1
        Call SE.module("DB").getDBConnection.Close()
        If SE.module("DB").getDBConnection.State = objectStateEnum.Item("adStateClosed") Then _
            close = 0
    End Function

    '''
     ' 执行SQL操作
     '
     ' @return recordset <数据集>
     ''
    Public Function executeSql(ByVal sqlString)
        If SE.module("DB").getDBConnection.State <> objectStateEnum.Item("adStateOpen") Then
            Set executeSql = Nothing
            Exit Function
        End If

        Set executeSql = SE.module("DB").getDBConnection.Execute(sqlString)
    End Function

End Class
%>