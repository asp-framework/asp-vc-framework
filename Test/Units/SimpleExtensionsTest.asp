<%
'''
 ' SimpleExtensionsTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.10.30
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- #include file = "../../Framework/SimpleExtensions.asp" -->

<%
Class SimpleExtensionsTest

    Private vActual

	Public Function TestCaseNames()
		TestCaseNames = Array( _
            "loadFileTest", _
            "getIncludeCodeTest", _
            "getIncludeHtmlTest" _
        )
	End Function

	Public Sub SetUp()
        ' Response.Write("SetUp<br>")
	End Sub

	Public Sub TearDown()
		'Response.Write("TearDown<br>")
	End Sub

    ' 读取文件测试
	Public Sub loadFileTest(oTestResult)
        vActual = SE.loadFile("./UserFiles/loadFileTest.asp")

		oTestResult.AssertEquals "读取文件测试", vActual, "读取文件信息异常"
	End Sub

    ' 包含并运行文件测试
    Public Sub getIncludeCodeTest(oTestResult)
        Response.Flush
        vActual = SE.getIncludeCode("./UserFiles/includeTest/includeTest1.asp")
        Response.Clear

		oTestResult.AssertEquals _
            "Response.Write(""开始文件导入测试<br/>"" & vbCrLf & """")" & vbCrLf _
            & "Dim output : output = ""成功输出内容""" & vbCrLf _
            & "Response.Write("""" & vbCrLf & """")" & vbCrLf _
            & "Response.Write(""output:""&output)" & vbCrLf _
            & "Response.Write("""")" & vbCrLf, _
            vActual,_
            "包含文件异常"
    End Sub

    ' 包含文件获取执行后的内容测试
    Public Sub getIncludeHtmlTest(oTestResult)
        vActual = SE.getIncludeHtml("./UserFiles/includeTest/includeTest1.asp")

        oTestResult.AssertEquals _
            "开始文件导入测试<br/>" & vbCrLf & vbCrLf & "output:成功输出内容", _
            vActual, _
            "载入配置文件异常"
    End Sub

End Class
%>