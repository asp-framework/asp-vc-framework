<%
'''
 ' SimpleExtensions.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.9.26
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- #include file = "SimpleExtensionsBase.asp" -->

<% Dim SE : Set SE = New SimpleExtensions %>
<%
Class SimpleExtensions

    ' @var class simpleExtensionsBaseClass <SimpleExtensionsBase类>
    Private simpleExtensionsBaseClass
    '''
     ' 获取 SimpleExtensionsBase 类
     ''
    Private Property Get getSimpleExtensionsBaseClass()
        If VarType(simpleExtensionsBaseClass) <> 9 Then Set simpleExtensionsBaseClass = New SimpleExtensionsBase
        Set getSimpleExtensionsBaseClass = simpleExtensionsBaseClass
    End Property

'###########################'
'###########################'

    '''
     ' 构造函数
     ''
    Private Sub Class_Initialize
        
    End Sub

'###########################'
'###########################'

    '''
     ' 运行框架
     ''
    Public Function run()
        ' 运行配置文件

        ' 运行路由

        ' 运行控制器

        ' 渲染视图

    End Function

'###########################'
'###########################'

    '''
     ' 读取文件
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <文件内容>
     ''
    Public Function loadFile(ByVal filePath)
        loadFile = getSimpleExtensionsBaseClass.loadFile(filePath)
    End Function

    '''
     ' 包含并运行文件
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <可执行代码>
     ''
    Public Function include(ByVal filePath)
        include = getSimpleExtensionsBaseClass.include(filePath)
    End Function

End Class
%>