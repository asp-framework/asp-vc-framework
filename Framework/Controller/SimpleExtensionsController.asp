<%
'''
 ' SimpleExtensionsController.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.31
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsController

    ' @var string <控制器路径>
    Private controllerPath

'###########################'
'###########################'

    '''
     '  运行控制器
     ''
    Public Function run()
        setControllerPath()
        SE.include(controllerPath)
        runAction()
    End Function

    '''
     '  设置控制器路径
     ''
    Private Function setControllerPath()
        controllerPath = _
            SE.getConfigs("system/appsDir/Value") _
            & "/" & SE.module("Router").getAppName _
            & "/" & "Controllers" _
            & "/" & SE.module("Router").getControllerName & "Controller" & ".asp"
    End Function

    '''
     '  运行动作
     ''
    Private Function runAction()
        Dim controller
        Set controller = Eval("New " & SE.module("Router").getControllerName & "Controller")
        Execute("controller." & SE.module("Router").getActionName & "Action()")
    End Function

End Class
%>