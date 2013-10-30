<%
'''
 ' SimpleExtensionsBaseTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.28
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
            "getIncludeHtmlTest", _
            "loadConfigsTest", _
            "moduleTest" _
        )
	End Function

	Public Sub SetUp()
		Set SimpleExtensionsBaseClass = New SimpleExtensionsBase
	End Sub

	Public Sub TearDown()
		'Response.Write("TearDown<br>")
	End Sub

    ' 读取文件测试
	Public Sub loadFileTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.loadFile("./UserFiles/loadFileTest.asp")

		oTestResult.AssertEquals _
            "读取文件测试", _
            vActual, _
            "读取文件信息异常"
	End Sub

    ' 包含文件获取可执行代码测试
    Public Sub getIncludeCodeTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.getIncludeCode("./UserFiles/includeTest/includeTest1.asp")

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
    Public Sub getIncludeHtmlTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.getIncludeHtml("./UserFiles/includeTest/includeTest1.asp")

        oTestResult.AssertEquals _
            "开始文件导入测试<br/>" & vbCrLf & vbCrLf & "output:成功输出内容", _
            vActual, _
            "载入配置文件异常"
    End Sub

    ' 载入配置文件测试
    Public Sub loadConfigsTest(oTestResult)
        SimpleExtensionsBaseClass.loadConfigs("./UserFiles/config.xml")

        oTestResult.AssertEquals _
            "../Framework", _
            SimpleExtensionsBaseClass.getConfigs(Null).Item("system").Item("seDir").Item("Value"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "../Framework", _
            SimpleExtensionsBaseClass.getConfigs("system/seDir/Value"), _
            "载入配置文件异常"

		oTestResult.AssertEquals _
            "HelloWorld", _
            SimpleExtensionsBaseClass.getConfigs(Null).Item("router").Item("defaultAppName").Item("Value"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "HelloWorld", _
            SimpleExtensionsBaseClass.getConfigs("router/defaultAppName/Value"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "get", _
            SimpleExtensionsBaseClass.getConfigs(Null).Item("router").Item("Attributes").Item("type"), _
            "载入配置文件异常"

        oTestResult.AssertEquals _
            "get", _
            SimpleExtensionsBaseClass.getConfigs("router/Attributes/type"), _
            "载入配置文件异常"
    End Sub

    ' 导入模块测试
    Public Sub moduleTest(oTestResult)
        SimpleExtensionsBaseClass.loadConfigs("./UserFiles/config.xml")
        vActual = SimpleExtensionsBaseClass.module("String").md5("SE")

        oTestResult.AssertEquals _
            "f003c44deab679aa2edfaff864c77402", _
            vActual, _
            "导入模块异常"
    End Sub

End Class
%>