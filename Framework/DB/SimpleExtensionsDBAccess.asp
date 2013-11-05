<%
'''
 ' SimpleExtensionsDBAccess.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.5
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
        If SE.module("DB").getDBConnection.State = objectStateEnum.Item("adStateClosed") Then _
        Call SE.module("DB").getDBConnection.Open( _
            "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & SE.module("DB").getDBSource & ";" & _
            "User Id=;" & _
            "Password=;" _
        )
    End Function

    '''
     ' 关闭数据库
     ''
    Public Function close()
        If SE.module("DB").getDBConnection.State <> objectStateEnum.Item("adStateClosed") Then _
        Call SE.module("DB").getDBConnection.Close()
    End Function

    '''
     ' 执行SQL操作
     ''
    Public Function sqlExecute(ByVal sqlString)
        If SE.module("DB").getDBConnection.State <> objectStateEnum.Item("adStateOpen") Then
            Set sqlExecute = Nothing
            Exit Function
        End If

        Set sqlExecute = SE.module("DB").getDBConnection.Execute(sqlString)
    End Function

End Class
%>