<%
'''
 ' SimpleExtensions.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.30
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- #include file = "SimpleExtensionsBase.asp" -->

<% Dim SE : Set SE = New SimpleExtensions %>
<%
Class SimpleExtensions

    ' @var class simpleExtensionsBaseClass <SE框架基类>
    Private simpleExtensionsBaseClass

'###########################'
'###########################'

    '''
     ' 获取SE框架基类
     ''
    Public Property Get getSimpleExtensionsBaseClass()
        If VarType(simpleExtensionsBaseClass) <> 9 Then Set simpleExtensionsBaseClass = New SimpleExtensionsBase
        Set getSimpleExtensionsBaseClass = simpleExtensionsBaseClass
    End Property

    '''
     ' 构造函数
     ''
    Private Sub Class_Initialize

    End Sub

'###########################'
'###########################'

    '''
     ' 运行框架
     '
     ' @param string configFilePath <配置文件路径>
     ''
    Public Function run(ByVal configFilePath)
        ' 运行配置文件
        getSimpleExtensionsBaseClass.loadConfigs(configFilePath)
        ' 运行路由

        ' 运行控制器

        ' 渲染视图

    End Function

'###########################'
'###########################'

    '''
     ' 获取配置项
     '
     ' @param string configPath <配置路径,例:"system/seDir/Value">
     '
     ' @return dictionary|string <所有配置数据|配置项字符串>
     ''
    Public Property Get getConfigs(ByVal configPath)
        Set getConfigs = getSimpleExtensionsBaseClass.getConfigs(configPath)
    End Property

    '''
     ' 获取框架根目录
     ''
    Public Property Get getSEDir()
        getSEDir = getSimpleExtensionsBaseClass.getSEDir
    End Property

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
     ' 包含并执行文件
     '
     ' @param string filePath <文件路径>
     ''
    Public Function include(ByVal filePath)
        getSimpleExtensionsBaseClass.include(filePath)
    End Function

    '''
     ' 包含并运行文件
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <可执行代码>
     ''
    Public Function getIncludeCode(ByVal filePath)
        getIncludeCode = getSimpleExtensionsBaseClass.getIncludeCode(filePath)
    End Function

    '''
     ' 包含文件获取执行后的内容,但不输出内容
     '
     ' @param string filePath <文件路径>
     '
     ' @return string <执行后的内容>
     ''
    Public Function getIncludeHtml(ByVal filePath)
        getIncludeHtml = getSimpleExtensionsBaseClass.getIncludeHtml(filePath)
    End Function

    '''
     ' 调用模块
     '
     ' @param string moduleName <模块名称>
     ''
    Public Function module(ByVal moduleName)
        Set module = getSimpleExtensionsBaseClass.module(moduleName)
    End Function

End Class
%>