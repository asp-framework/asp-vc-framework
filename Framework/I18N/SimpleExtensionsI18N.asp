<%
'''
 ' SimpleExtensionsI18N.asp 文件
 ' @author 高翔 <263027768@qq.com>
 ' @version 2013.11.3
 ' @copyright Copyright (c) 2013-2014 SE
 ''
%>

<%
Class SimpleExtensionsI18N

    ' @var string <当前语言>
    Private localLanguage

    ' @var dictionary <翻译的内容>
    Private tContent

'###########################'
'###########################'

    '''
     ' 构造函数
     ''
    Private Sub Class_Initialize
        setLocalLanguage(SE.getConfigs("I18N/language/Value"))
    End Sub

    '''
     '  翻译指定信息到当前语言
     ''
    Public Function t(ByVal keyPath)

    End Function

'###########################'
'###########################'

    '''
     '  设置当前语言
     ''
     Public Function setLocalLanguage(ByVal languageString)
        loadTContent(languageString)
        localLanguage = languageString
     End Function

    '''
     '  读取翻译内容
     ''
     Private Function loadTContent(ByVal languageString)

     End Function

    '''
     '  获取当前语言
     ''
    Public Property Get getLocalLanguage()
        getLanguage = language
    End Property

End Class
%>