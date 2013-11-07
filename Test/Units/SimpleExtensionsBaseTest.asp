<%
'''
 ' SimpleExtensionsBaseTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.30
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsBaseTest

    Private SimpleExtensionsBaseClass

    Private vActual

	Public Function TestCaseNames()
		TestCaseNames = Array( _
            "loadFileTest", _
            "getIncludeCodeTest", _
            "getIncludeResultTest", _
            "loadConfigsTest", _
            "getSEDirTest", _
            "isDevelopmentTest", _
            "moduleTest" _
        )
	End Function

	Public Sub SetUp()
		Set SimpleExtensionsBaseClass = New SimpleExtensionsBase
        SimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
	End Sub

	Public Sub TearDown()
		'Response.Write("TearDown<br>")
	End Sub

    ' 读取文件测试
	Public Sub loadFileTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.loadFile("./ProjectTest/IncludeTest/loadFileTest.asp")

		oTestResult.AssertEquals _
            "读取文件测试", _
            vActual, _
            "读取文件信息异常"
	End Sub

    ' 包含文件获取可执行代码测试
    Public Sub getIncludeCodeTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.getIncludeCode("./ProjectTest/IncludeTest/IncludeTest/includeTest1.asp")

		oTestResult.AssertEquals _
            "Response.Write(""开始文件导入测试<br/>"" & vbCrLf & """")" & vbCrLf _
            & "Dim output : output = ""成功输出内容""" & vbCrLf _
            & "Response.Write("""" & vbCrLf & """")" & vbCrLf _
            & "Response.Write(""output:""&output)" & vbCrLf _
            & "Response.Write("""")" & vbCrLf, _
            vActual, _
            "包含文件异常"
    End Sub

    ' 包含文件获取执行后的内容测试
    Public Sub getIncludeResultTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.getIncludeResult("./ProjectTest/IncludeTest/IncludeTest/includeTest1.asp")

        oTestResult.AssertEquals _
            "开始文件导入测试<br/>" & vbCrLf & vbCrLf & "output:成功输出内容", _
            vActual, _
            "载入配置文件异常"
    End Sub

    ' 载入配置文件测试
    Public Sub loadConfigsTest(oTestResult)
        oTestResult.AssertEquals _
            "../Framework", _
            SimpleExtensionsBaseClass.getConfigs(Null).Item("System").Item("seDir").Item("Value"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "../Framework", _
            SimpleExtensionsBaseClass.getConfigs("System/seDir/Value"), _
            "载入配置文件异常"

		oTestResult.AssertEquals _
            "Test", _
            SimpleExtensionsBaseClass.getConfigs(Null).Item("Router").Item("appName").Item("Value"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "Test", _
            SimpleExtensionsBaseClass.getConfigs("Router/appName/Value"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "get", _
            SimpleExtensionsBaseClass.getConfigs(Null).Item("Router").Item("Attributes").Item("type"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "get", _
            SimpleExtensionsBaseClass.getConfigs("Router/Attributes/type"), _
            "载入配置文件异常"
    End Sub

    ' 获取框架根目录测试
    Public Sub getSEDirTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.getSEDir

        oTestResult.AssertEquals _
            "../Framework", _
            vActual, _
            "获取框架根目录异常"
    End Sub

    ' 判断是否开发环境测试
    Public Sub isDevelopmentTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.isDevelopment

        oTestResult.AssertEquals _
            True, _
            vActual, _
            "判断是否开发环境异常"
    End Sub

    ' 导入模块测试
    Public Sub moduleTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.module("String").md5("SE")

        oTestResult.AssertEquals _
            "f003c44deab679aa2edfaff864c77402", _
            vActual, _
            "导入模块异常"
    End Sub

End Class
%>