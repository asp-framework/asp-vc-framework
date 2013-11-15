getLocalLanguage
================
`getLocalLanguage` &mdash; 获取当前语言

说明
----
>     string getLocalLanguage(void)
> 获取当前语言。

参数
----
> 没有参数。

返回值
------
> **类型：**`string`  
> **说明：**当前语言。

范例
----
>
    <%
    Dim language
    language = SE.module("I18N").getLocalLanguage
    %>