<%
'''
 ' SimpleExtensionsDBTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.6
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsDBTest

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "openTest", _
            "closeTest", _
            "executeSqlTest", _
            "commandTest" _
        )
    End Function

    Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 打开数据库测试
    Public Sub openTest(oTestResult)
        SE.module("DB").open()
        vActual = SE.module("DB").getDBConnection.State

        oTestResult.AssertEquals _
            1, _
            vActual, _
            "打开数据库异常"
    End Sub

    ' 关闭数据库测试
    Public Sub closeTest(oTestResult)
        SE.module("DB").close()
        vActual = SE.module("DB").getDBConnection.State

        oTestResult.AssertEquals _
            0, _
            vActual, _
            "打开数据库异常"
    End Sub

    ' 执行SQL操作测试
    Public Sub executeSqlTest(oTestResult)
        SE.module("DB").open()
        Set vActual = SE.module("DB").executeSql("SELECT userName FROM UserLists")

        oTestResult.AssertEquals _
            "Admin", _
            vActual.Fields("userName"), _
            "打开数据库异常"

        SE.module("DB").close()
    End Sub

    ' 命令测试
    Public Sub commandTest(oTestResult)
        SE.module("DB").open()

        SE.module("DB").command.createCommand("SELECT userName FROM UserLists WHERE userName = :userName")
        Call SE.module("DB").command.bindParameter(":userName", "Admin", "dbString")
        Set vActual = SE.module("DB").command.executeCommand()

        oTestResult.AssertEquals _
            "Admin", _
            vActual.Fields("userName"), _
            "命令异常"

        SE.module("DB").close()
    End Sub

End Class
%>