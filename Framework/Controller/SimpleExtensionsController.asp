<%
'''
 ' SimpleExtensionsController.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.17
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsController

    ' @var string <当前应用目录>
    ' 获取函数: getAppDir
    Private appDir

    ' @var string <当前应用控制器目录>
    ' 获取函数: getControllersDir
    Private controllersDir

    ' @var string <当前控制器视图目录>
    ' 获取函数: getViewsDir
    Private viewsDir

    ' @var string <当前控制器名称>
    ' 获取函数: getControllerName
    Private controllerName

    ' @var dictionary controllersQueue <控制器队列>
    Private controllersQueue

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
    End Sub

    '''
     ' 运行控制器
     ''
    Public Function run()
        checkError()
        Call runAction(Me.getControllerName, SE.module("Router").getActionName)
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
        Dim controllerPath
        controllerPath = _
            getControllersDir _
            & "/" & SE.module("Router").getControllerName & "Controller" & ".asp"
        If Not SE.module("File").fileExists(controllerPath) Then _
            Call SE.module("Error").throwError( _
                2, _
                "控制器【" & Me.getControllerName & "】不存在。" _
            )
    End Function

    '''
     ' 运行动作
     '
     ' @param string controllerName <控制器名称>
     ' @param string actionName <动作名称>
     ''
    Public Function runAction(ByVal controllerName, ByVal actionName)
        Call runFunction(controllerName, actionName & "Action", Null)
    End Function

    '''
     ' 运行方法
     '
     ' @param string controllerName <控制器名称>
     ' @param string functionName <方法名称>
     ' @param array|null parameters <方法需要的参数>
     ''
    Public Function runFunction(ByVal controllerName, ByVal functionName, ByVal parameters)
        If VarType(controllersQueue) <> 9 Then _
            Set controllersQueue = Server.CreateObject("Scripting.Dictionary")

        ' 向队列添加控制器
        If Not controllersQueue.Exists(controllerName) Then
            SE.include(getControllersDir & "/" & controllerName & "Controller.asp")
            Call controllersQueue.Add(controllerName, Eval("New " & controllerName & "Controller"))
        End If

        ' 方法需要的参数
        Dim functionParameters
        If IsArray(parameters) Then
            functionParameters = "parameters(0)"
            Dim parametersCounter
            For parametersCounter = 1 To UBound(parameters)
                functionParameters = functionParameters & ", parameters(" & CStr(parametersCounter) & ")"
            Next
        End If

        On Error Resume Next
        Execute("Call controllersQueue.Item(""" & controllerName & """)." & _
            functionName & "(" & functionParameters & ")")
        If Err.Number = 438 Then _
            Call SE.module("Error").throwError( _
                2, _
                "方法【" & functionName & "】不存在。" _
            )
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