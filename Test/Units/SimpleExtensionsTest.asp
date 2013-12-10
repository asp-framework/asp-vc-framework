<%
'''
 ' SimpleExtensionsTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.12.10
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsTest

    Private vActual

	Public Function TestCaseNames()
		TestCaseNames = Array( _
            "getConfigsTest", _
            "getSEDirTest", _
            "isDevelopmentTest", _
            "getIncludeCodeTest", _
            "getIncludeResultTest", _
            "moduleTest" _
        )
	End Function

	Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
	End Sub

	Public Sub TearDown()
		'Response.Write("TearDown<br>")
	End Sub

    ' 获取配置项测试
    Public Sub getConfigsTest(oTestResult)
        oTestResult.AssertEquals _
            "../Framework", _
            SE.getConfigs(Null).Item("System").Item("seDir").Item("Value"), _
            "获取配置项异常"

        oTestResult.AssertEquals _
            "../Framework", _
            SE.getConfigs("System/seDir/Value"), _
            "获取配置项异常"

        oTestResult.AssertEquals _
            "Test", _
            SE.getConfigs(Null).Item("Router").Item("appName").Item("Value"), _
            "获取配置项异常"

        oTestResult.AssertEquals _
            "Test", _
            SE.getConfigs("Router/appName/Value"), _
            "获取配置项异常"

        oTestResult.AssertEquals _
            "get", _
            SE.getConfigs(Null).Item("Router").Item("Attributes").Item("type"), _
            "获取配置项异常"

        oTestResult.AssertEquals _
            "get", _
            SE.getConfigs("Router/Attributes/type"), _
            "获取配置项异常"
    End Sub

    ' 获取框架根目录测试
    Public Sub getSEDirTest(oTestResult)
        vActual = SE.getSEDir

        oTestResult.AssertEquals _
            "../Framework", _
            vActual, _
            "获取框架根目录异常"
    End Sub

    ' 判断是否开发环境测试
    Public Sub isDevelopmentTest(oTestResult)
        vActual = SE.isDevelopment

        oTestResult.AssertEquals _
            True, _
            vActual, _
            "判断是否开发环境异常"
    End Sub

    ' 包含并运行文件测试
    Public Sub getIncludeCodeTest(oTestResult)
        vActual = SE.getIncludeCode("./ProjectTest/includeTest/IncludeTest/includeTest1.asp")

		oTestResult.AssertEquals _
            "Response.Write(""开始文件导入测试<br/>"" & vbCrLf & """")" & vbCrLf _
            & "Dim output : output = ""成功输出内容""" & vbCrLf _
            & "Response.Write("""" & vbCrLf & """")" & vbCrLf _
            & "Response.Write(""output:""&output)" & vbCrLf, _
            vActual,_
            "包含文件异常"
    End Sub

    ' 包含文件获取执行后的内容测试
    Public Sub getIncludeResultTest(oTestResult)
        vActual = SE.getIncludeResult("./ProjectTest/IncludeTest/IncludeTest/includeTest1.asp")

        oTestResult.AssertEquals _
            "开始文件导入测试<br/>" & vbCrLf & vbCrLf & "output:成功输出内容", _
            vActual, _
            "载入配置文件异常"
    End Sub

    ' 导入模块测试
    Public Sub moduleTest(oTestResult)
        vActual = SE.module("String").md5("SE")

        oTestResult.AssertEquals _
            "f003c44deab679aa2edfaff864c77402", _
            vActual, _
            "导入模块异常"
    End Sub

End Class
%>