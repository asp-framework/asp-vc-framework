<%
'''
 ' SimpleExtensionsBaseTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.9.26
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsBaseTest

    Private SimpleExtensionsBaseClass

    Private vActual

	Public Function TestCaseNames()
		TestCaseNames = Array(_
            "aspIncludeTagTest",_
            "loadFileTest",_
            "includeTest",_
            "loadConfigsTest"_
        )
	End Function

	Public Sub SetUp()
		Set SimpleExtensionsBaseClass = New SimpleExtensionsBase
	End Sub

	Public Sub TearDown()
		'Response.Write("TearDown<br>")
	End Sub

    ' ASP #include 标签是否开启测试
    Public Sub aspIncludeTagTest(oTestResult)
        SimpleExtensionsBaseClass.setAspIncludeTag = False
        vActual = SimpleExtensionsBaseClass.isAspIncludeTag

        oTestResult.AssertEquals False, vActual, "ASP #include 标签关闭失败"
    End Sub

    ' 读取文件测试
	Public Sub loadFileTest(oTestResult)
        vActual = SimpleExtensionsBaseClass.loadFile("./UserFiles/loadFileTest.asp")

		oTestResult.AssertEquals "读取文件测试", vActual, "读取文件信息异常"
	End Sub

    ' 包含并运行文件测试
    Public Sub pressModeIncludeTest(oTestResult)
        Response.Flush
        vActual = SimpleExtensionsBaseClass.pressModeInclude("./UserFiles/includeTest/includeTest1.asp")
        Response.Clear

		oTestResult.AssertEquals _
            "Response.Write(""开始文件导入测试<br/>"" & vbCrLf & """")" & vbCrLf _
            & "Dim output : output = ""成功输出内容""" & vbCrLf _
            & "Response.Write("""" & vbCrLf & """")" & vbCrLf _
            & "Response.Write(""output:""&output)" & vbCrLf _
            & "Response.Write("""")" & vbCrLf,_
            vActual,_
            "包含文件异常"
    End Sub

    ' 载入配置文件
    Public Sub loadConfigsTest(oTestResult)
        SimpleExtensionsBaseClass.loadConfigs("./UserFiles/config.xml")

		oTestResult.AssertEquals _
            True,_
            True,_
            "载入配置文件异常"
    End Sub

End Class
%>