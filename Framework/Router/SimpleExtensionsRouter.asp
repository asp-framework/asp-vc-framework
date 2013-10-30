<%
'''
 ' SimpleExtensionsRouter.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.30
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsRouter

    ' @var string <应用名称>
    Private appName

    ' @var string <控制器名称>
    Private controllerName

    ' @var string <动作名称>
    Private actionName

'###########################'
'###########################'

    '''
     '  运行路由
     ''
    Public Function run()

    End Function

    Public Function getRequestValue(ByVal variable, ByVal index)
        If Not IsNumeric(index) Or index < 1 Then index = 1
        getRequestValue = Request.QueryString(variable)(index)
    End Function

'###########################'
'###########################'

    '''
     ' 获取应用名称
     ''
    Public Property Get getAppName()
        getApp = app
    End Property

    '''
     ' 获取控制器名称
     ''
    Public Property Get getControllerName()
        getController = controller
    End Property

    '''
     ' 获取动作名称
     ''
    Public Property Get getActionName()
        getAction = action
    End Property

End Class
%>