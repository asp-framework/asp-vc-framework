<%
'''
 ' SimpleExtensionsDBCommand.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.6
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsDBCommand

    ' @var string <命令>
    Private commandString

    ' @var dictionary <绑定的参数列表>
    Private bindParameterList

    ' @var dictionary <数据类型>
    Private dataTypeList

'###########################'
'###########################'

    Private Sub Class_Initialize
        Set bindParameterList = Server.CreateObject("Scripting.Dictionary")
        initDataTypeList()
    End Sub

    '''
     ' 初始化数据类型列表
     ''
    Private Sub initDataTypeList()
        Set dataTypeList = Server.CreateObject("Scripting.Dictionary")
        Call dataTypeList.Add("dbString", 0)
        Call dataTypeList.Add("dbInteger", 1)
    End Sub

    '''
     ' 创建命令
     '
     ' @param string sqlString <SQL命令字符串>
     ''
    Public Function createCommand(ByVal sqlString)
        commandString = sqlString
    End Function

    '''
     ' 绑定参数
     '
     ' @param string name <绑定的参数名>
     ' @param mixed value <绑定的参数值>
     ' @param string dataType <绑定参数的类型>
     '
     ' @return boolean <是否绑定成功>
     ''
    Public Function bindParameter(ByVal name, ByVal value, ByVal dataType)
        bindParameter = processParameterToSafe(value, dataType)

        If Not bindParameter Then Exit Function

        If bindParameterList.Exists(name) Then
            bindParameterList.Item(name) = value
        Else
            Call bindParameterList.Add(name, value)
        End If
    End Function

    '''
     ' 处理绑定参数为安全参数
     '
     ' @param mixed value <绑定的值>
     ' @param string dataType <绑定值的类型,参照[dataTypeList]>
     '
     ' @return boolean <处理参数是否成功>
     ''
    Private Function processParameterToSafe(ByRef value, ByVal dataType)
        processParameterToSafe = True
        If Not dataTypeList.Exists(dataType) Then
            processParameterToSafe = False
            Exit Function
        End If
        dataType = dataTypeList.Item(dataType)
        Select Case dataType
            ' dbString
            Case 0
                value = Replace(value, "'", "''")
                value = "'" & value & "'"
                value = CStr(value)
            ' dbInteger
            Case 1
                If IsNumeric(value) Then
                    value = CInt(value)
                Else
                    processParameterToSafe = False
                End If
        End Select
    End Function

    '''
     ' 执行命令
     '
     ' @return recordset <数据集>
     ''
    Public Function executeCommand()
        parseBindParamerters()
        Set executeCommand = SE.module("DB").executeSql(commandString)
    End Function

    '''
     ' 解析绑定参数
     ''
    Private Function parseBindParamerters()
        Dim keysArray, nowKey
        keysArray = bindParameterList.Keys
        For Each nowKey In keysArray
            commandString = Replace(commandString, nowKey, bindParameterList.Item(nowKey))
        Next
    End Function

End Class
%>