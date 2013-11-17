md5
===
`md5` &mdash; MD5加密

说明
----
>     string md5(string stringToMD5)
> MD5加密。

参数
----
> `stringToMD5`
>> **类型：** `string`  
>> **说明：**需要加密到MD5的字符串。

返回值
------
> **类型：**`string`  
> **说明：**MD5字符串。

范例
----
>
    <%
    Dim stringToMD5, md5String
    stringToMD5 = "SE"
    md5String = SE.module("String").md5(stringToMD5)
    Response.Write(md5String)
    %>
> 以上内容输出：
>
    f003c44deab679aa2edfaff864c77402