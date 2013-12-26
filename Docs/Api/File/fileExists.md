fileExists
==========
`fileExists` &mdash; 判断文件是否存在

说明
----
>     boolean fileExists(string filePath)
> 判断文件是否存在。

参数
----
> `filePath`
>> **类型：**`string`  
>> **说明：**文件路径。(路径格式为相对路径，相对路径起始于首个调用文件目录。)

返回值
------
> **类型：**`boolean`  
> **说明：**文件是否存在。

范例
----
>
    <%
    Dim filePath
    filePath = "./path/to/fileName"
    If SE.module("File").fileExists(filePath) Then
        Response.Write("文件存在。")
    Else
        Response.Write("文件不存在。")
    End If
    %>