<%
'''
 ' SimpleExtensionsDB.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.6
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- 导入支持文件 -->
    <!-- #include file = "SimpleExtensionsDBCommand.asp" -->
    <!-- #include file = "SimpleExtensionsDBAccess.asp" -->
<!-- /导入支持文件 -->

<%
Class SimpleExtensionsDB

    ' @var string <数据库类型>
    Private dbType

    ' @var string <数据库源>
    Private dbSource

    ' @var string <数据库名称>
    Private dbName

    ' @var string <数据库用户名>
    Private dbUserName

    ' @var string <数据库密码>
    Private dbPassword

    ' @var object <数据库连接>
    Private dbConnection

    ' @var class <数据库解析类>
    Private dbParseClassByType

    ' @var class <命令类>
    Private commandClass

'###########################'
'###########################'

    Private Sub Class_Initialize
        initConfigs()

        ' 初始化数据库连接
        Set dbConnection = Server.CreateObject("ADODB.Connection")
        ' 初始化当前数据库类型处理类
        Execute("Set dbParseClassByType = " & "New SimpleExtensionsDB" & dbType)
    End Sub

    '''
     ' 初始化配置项
     ''
    Private Sub initConfigs()
        dbType = SE.getConfigs("DB/type/Value")
        dbSource = SE.getConfigs("DB/source/Value")
        dbName = SE.getConfigs("DB/dbName/Value")
        dbUserName = SE.getConfigs("DB/userName/Value")
        dbPassword = SE.getConfigs("DB/password/Value")
    End Sub

    '''
     ' 打开数据库
     ''
    Public Function open()
        open = Execute("dbParseClassByType." & "open()" )
    End Function

    '''
     ' 关闭数据库
     ''
    Public Function close()
       close = Eval("dbParseClassByType." & "close()" )
    End Function

    '''
     ' 执行SQL操作
     '
     ' @return recordset <数据集>
     ''
    Public Function executeSql(ByVal sqlString)
        Set executeSql = Eval("dbParseClassByType." & "executeSql(sqlString)" )
    End Function

    '''
     ' 命令对象
     '
     ' @return class <命令类>
     ''
    Public Property Get command()
        If VarType(commandClass) <> 9 Then _
            Set commandClass = New SimpleExtensionsDBCommand
        Set command = commandClass
    End Property

'###########################'
'###########################'

    '''
     ' 获取数据库类型
     ''
    Public Property Get getDBType()
        getDBType = dbType
    End Property

    '''
     ' 获取数据库源
     ''
    Public Property Get getDBSource()
        getDBSource = dbSource
    End Property

    '''
     ' 获取数据库名称
     ''
    Public Property Get getDBName()
        getDBName = dbName
    End Property

    '''
     ' 获取数据库用户名
     ''
    Public Property Get getDBUserName()
        getDBUserName = dbUserName
    End Property

    '''
     ' 获取数据库密码
     ''
    Public Property Get getDBPassword()
        getDBPassword = dbPassword
    End Property

    '''
     ' 获取数据库连接
     ''
    Public Property Get getDBConnection()
        Set getDBConnection = dbConnection
    End Property

End Class
%>