<%
'''
 ' SimpleExtensionsI18NTest.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.3
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsI18NTest

    Private vActual

    Public Function TestCaseNames()
        TestCaseNames = Array( _
            "setAndGetLocalLanguageTest", _
            "tTest" _
        )
    End Function

    Public Sub SetUp()
        SE.getSimpleExtensionsBaseClass.loadConfigs("./ProjectTest/Configs/config.xml")
        SE.module("Router").run()
    End Sub

    Public Sub TearDown()
        'Response.Write("TearDown<br>")
    End Sub

    ' 本地语言设置获取测试
    Public Sub setAndGetLocalLanguageTest(oTestResult)
        vActual = SE.module("I18N").getLocalLanguage

        oTestResult.AssertEquals _
            "zh-cn", _
            vActual, _
            "读取文件信息异常"

        SE.module("I18N").setLocalLanguage("en-us")
        vActual = SE.module("I18N").getLocalLanguage

        oTestResult.AssertEquals _
            "en-us", _
            vActual, _
            "读取文件信息异常"
    End Sub

    ' 翻译测试
    Public Sub tTest(oTestResult)
        SE.module("I18N").setLocalLanguage("zh-cn")
        vActual = SE.module("I18N").t("Body/content/Value")

        oTestResult.AssertEquals _
            "你好～！", _
            vActual, _
            "读取文件信息异常"

        SE.module("I18N").setLocalLanguage("en-us")
        vActual = SE.module("I18N").t("Body/content/Value")

        oTestResult.AssertEquals _
            "Hello World~!", _
            vActual, _
            "读取文件信息异常"
    End Sub

End Class
%>