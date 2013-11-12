<%
'''
 ' SimpleExtensionsController.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.7
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsController

    ' @var string <当前应用目录>
    Private appDir

    ' @var string <当前应用控制器目录>
    Private controllersDir

    ' @var string <当前控制器视图目录>
    Private viewsDir

    ' @var string <控制器路径>
    Private controllerPath

    ' @var string <当前控制器名称>
    Private controllerName

'###########################'
'###########################'

    Private Sub Class_Initialize
        initConfigs()
    End Sub

    '''
     ' 初始化配置项
     ''
    Private Sub initConfigs()
        appDir = SE.getConfigs("System/appsDir/Value") & "/" & SE.module("Router").getAppName
        controllersDir = getAppDir & "/Controllers"
        controllerName = SE.module("Router").getControllerName
        viewsDir = getAppDir & "/Views/" & getControllerName
        controllerPath = _
            getControllersDir _
            & "/" & SE.module("Router").getControllerName & "Controller" & ".asp"
    End Sub

    '''
     ' 运行控制器
     ''
    Public Function run()
        checkError()
        SE.include(controllerPath)
        runAction()
    End Function

    '''
     ' 错误验证
     ''
    Private Function checkError()
        ' 判断应用是否存在
        If Not SE.module("file").dirExists(getAppDir) Then _
            Call SE.module("Error").throwError( _
                2, _
                "应用【" & SE.module("Router").getAppName & "】不存在。" _
            )

        ' 判断控制器是否存在
        If Not SE.module("File").fileExists(controllerPath) Then _
            Call SE.module("Error").throwError( _
                2, _
                "控制器【" & Me.getControllerName & "】不存在。" _
            )
    End Function

    '''
     ' 运行动作
     ''
    Private Function runAction()
        Dim controller
        Set controller = Eval("New " & getControllerName & "Controller")
        On Error Resume Next
        Execute("controller." & SE.module("Router").getActionName & "Action()")
        If Err.Number = 438 Then
            Call SE.module("Error").throwError( _
                2, _
                "动作【" & SE.module("Router").getActionName & "】不存在。" _
            )
            Err.Clear
        End If
        On Error GoTo 0
    End Function

'###########################'
'###########################'

    '''
     ' 获取布局路径
     '
     ' @param string layoutName <布局名称>
     '
     ' @return string <布局路径>
     ''
    Public Property Get getLayoutPath(ByVal layoutName)
        getLayoutPath = getAppDir & "/Views/Layouts/" & layoutName & ".asp"
    End Property

    '''
     ' 获取当前应用目录
     '
     ' @return string <当前应用目录>
     ''
    Public Property Get getAppDir()
        getAppDir = appDir
    End Property

    '''
     ' 获取当前控制器视图路径
     '
     ' @param string viewName <视图名称>
     '
     ' @return string <视图路径>
     ''
    Public Property Get getViewPath(ByVal viewName)
        getViewPath = getViewsDir & "/" & viewName & ".asp"
    End Property

    '''
     ' 获取当前控制器的视图目录
     '
     ' @return string <当前控制器的视图目录>
     ''
    Public Property Get getViewsDir()
        getViewsDir = viewsDir
    End Property

    '''
     ' 获取当前应用的控制器目录
     '
     ' @return string <当前应用的控制器目录>
     ''
    Public Property Get getControllersDir()
        getControllersDir = controllersDir
    End Property

    '''
     ' 获取当前控制器名称
     '
     ' @return string <当前控制器名称>
     ''
    Public Property Get getControllerName()
        getControllerName = controllerName
    End Property

End Class
%>