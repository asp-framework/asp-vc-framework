<%
'''
 ' SimpleExtensionsTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.9.26
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<!-- #include file = "../../Framework/SimpleExtensions.asp" -->

<%
Class SimpleExtensionsTest

    Private vActual

	Public Function TestCaseNames()
		TestCaseNames = Array(_
            "loadFileTest",_
            "includeTest"_
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
    Public Sub includeTest(oTestResult)
        Response.Flush
        vActual = SE.include("./UserFiles/includeTest/includeTest1.asp")
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
End Class
%>