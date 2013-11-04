<%@
    Language = "VBScript"
    CodePage = "65001"
%>
<%
Option Explicit
%>
<%
    Session.CodePage = 65001
    Response.Charset = "UTF-8"

    Response.CacheControl = "no-cache"
    Call Response.AddHeader("Pragma", "no-cache")
    Response.Expires = -1
%>
<!-- #include file = "Include/ASPUnitRunner.asp" -->

<!-- 导入测试文件 -->
<!-- #include file = "Units/SimpleExtensionsBaseTest.asp" -->
<!-- #include file = "Units/SimpleExtensionsTest.asp" -->
<!-- #include file = "Units/SimpleExtensionsRouterTest.asp" -->
<!-- #include file = "Units/SimpleExtensionsControllerTest.asp" -->
<!-- #include file = "Units/SimpleExtensionsI18NTest.asp" -->

<%
	Dim oRunner
	Set oRunner = New UnitRunner

    ' 实例化需要测试的类
    oRunner.AddTestContainer New SimpleExtensionsBaseTest
    oRunner.AddTestContainer New SimpleExtensionsTest
    oRunner.AddTestContainer New SimpleExtensionsRouterTest
    oRunner.AddTestContainer New SimpleExtensionsControllerTest
    oRunner.AddTestContainer New SimpleExtensionsI18NTest

	oRunner.Display()
%>
