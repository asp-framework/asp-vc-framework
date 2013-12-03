getUrl
======
`getUrl` &mdash; 获取URL

说明
----
>     string getUrl(string|integer urlType)

参数
----
> `urlType`
>> **类型：**`string` | `integer`  
>> **说明：**获取的URL类型。
>>> 目录：`"Dir"` 或 `0`  
>>> 路径：`"Path"` 或 `1`  
>>> 目录 + 询问参数：`"DirWith"` 或 `2`  
>>> 路径 + 询问参数：`"PathWith"` 或 `3`

返回值
------
> **类型：**`string` | `null`  
> **说明：**URL字符串。

范例
----
> 测试URL：[http://localhost/home/index.asp?a=1&b=2](#)
>
    <%
    Dim url1, url2
    url1 = SE.module("Request").getUrl("Dir")
    url2 = SE.module("Request").getUrl(0)
    Response.Write(url1 & "<br />" & url2 "<br />")
>
    url1 = SE.module("Request").getUrl("Path")
    url2 = SE.module("Request").getUrl(1)
    Response.Write(url1 & "<br />" & url2 "<br />")
>
    url1 = SE.module("Request").getUrl("DirWith")
    url2 = SE.module("Request").getUrl(2)
    Response.Write(url1 & "<br />" & url2 "<br />")
>
    url1 = SE.module("Request").getUrl("PathWith")
    url2 = SE.module("Request").getUrl(3)
    Response.Write(url1 & "<br />" & url2 "<br />")
    %>
> 以上内容输出：
>
    /test/
    /test/
    /test/index.asp
    /test/index.asp
    /test/?a=1&b=2
    /test/?a=1&b=2
    /test/index.asp?a=1&b=2
    /test/index.asp?a=1&b=2