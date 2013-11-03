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
    Private language

    ' @var dictionary <翻译的内容>
    Private tContent

'###########################'
'###########################'

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
     Public Function setLanguage(ByVal languageString)
        loadTContent(languageString)
        language = languageString
     End Function

    '''
     '  读取翻译内容
     ''
     Private Function loadTContent(ByVal languageString)

     End Function

    '''
     '  获取当前语言
     ''
    Public Property Get getLanguage()
        getLanguage = language
    End Property

End Class
%>