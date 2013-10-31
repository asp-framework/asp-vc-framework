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
        appName = getRequestValue("App")
        controllerName = getRequestValue("C")
        actionName = getRequestValue("A")
    End Function

    '''
     '  获取传入参数
     ''
    Private Function getRequestValue(ByVal variable)
        getRequestValue = Request.QueryString(variable)
        If Len(Request.Form(variable)) > 0 Then getRequestValue = Request.Form(variable)
        getRequestValue = requestValueToSafe(getRequestValue)
    End Function

    '''
     '  传入参数安全处理
     ''
    Private Function requestValueToSafe(ByVal toSafeValue)
        toSafeValue = Replace(toSafeValue, Space(1), Space(0))
        toSafeValue = Replace(toSafeValue, "@", Space(0))
        toSafeValue = Replace(toSafeValue, ":", Space(0))
        toSafeValue = Replace(toSafeValue, "_", Space(0))
        toSafeValue = Replace(toSafeValue, """", Space(0))
        toSafeValue = Replace(toSafeValue, "=", Space(0))
        requestValueToSafe = toSafeValue
    End Function

'###########################'
'###########################'

    '''
     ' 获取应用名称
     ''
    Public Property Get getAppName()
        getAppName = appName
    End Property

    '''
     ' 获取控制器名称
     ''
    Public Property Get getControllerName()
        getControllerName = controllerName
    End Property

    '''
     ' 获取动作名称
     ''
    Public Property Get getActionName()
        getActionName = actionName
    End Property

End Class
%>