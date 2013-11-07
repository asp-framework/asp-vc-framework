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

    ' @var object <数据库连接状态,
    '   0:关闭
    '   1:开启
    ' >
    Private dbConnectionStatus

    ' @var class <数据库解析类>
    Private dbParseClassByType

    ' @var class <命令类>
    Private commandClass

'###########################'
'###########################'

    Private Sub Class_Initialize
        initConfigs()
        checkConfigs()

        ' 初始化当前数据库类型处理类
        Set dbParseClassByType = Eval("New SimpleExtensionsDB" & dbType)
        ' 初始化数据库连接驱动
        Set dbConnection = dbParseClassByType.getConnectionDrive
        dbConnectionStatus = 0
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
     ' 验证基本配置项
     ''
    Private Function checkConfigs()
        If IsEmpty(dbType) Then _
            Call SE.module("Error").throwError( _
                2, _
                "请设置访问的数据库类型。" _
            )
    End Function

    '''
     ' 打开数据库
     ''
    Public Function open()
        dbParseClassByType.checkConfigs()

        dbConnectionStatus = dbParseClassByType.open()
        If dbConnectionStatus <> 1 Then _
            Call SE.module("Error").throwError( _
                2, _
                "数据库打开失败。" _
            )
    End Function

    '''
     ' 关闭数据库
     ''
    Public Function close()
       dbConnectionStatus = dbParseClassByType.close()
    End Function

    '''
     ' 执行SQL操作
     '
     ' @return recordset <数据集>
     ''
    Public Function executeSql(ByVal sqlString)
        Set executeSql = dbParseClassByType.executeSql(sqlString)
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
     '
     ' @return string <数据库类型>
     ''
    Public Property Get getDBType()
        getDBType = dbType
    End Property

    '''
     ' 获取数据库源
     '
     ' @return string <数据库源>
     ''
    Public Property Get getDBSource()
        getDBSource = dbSource
    End Property

    '''
     ' 获取数据库名称
     '
     ' @return string <数据库名称>
     ''
    Public Property Get getDBName()
        getDBName = dbName
    End Property

    '''
     ' 获取数据库用户名
     '
     ' @return string <数据库用户名>
     ''
    Public Property Get getDBUserName()
        getDBUserName = dbUserName
    End Property

    '''
     ' 获取数据库密码
     '
     ' @return string <数据库密码>
     ''
    Public Property Get getDBPassword()
        getDBPassword = dbPassword
    End Property

    '''
     ' 获取数据库连接
     '
     ' @return object <数据库连接>
     ''
    Public Property Get getDBConnection()
        Set getDBConnection = dbConnection
    End Property

    '''
     ' 获取数据库连接状态
     '
     ' @return integer <数据库连接状态,
     '   0:关闭
     '   1:开启
     ' >
     ''
    Public Property Get getDBConnectionStatus()
        getDBConnectionStatus = dbConnectionStatus
    End Property

End Class
%>