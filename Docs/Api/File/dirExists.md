dirExists
=========
`dirExists` &mdash; 判断目录是否存在

说明
----
>     boolean dirExists(string directory)
> 判断目录是否存在。

参数
----
> `directory`
>> **类型：**`string`  
>> **说明：**目录路径。(路径格式为相对路径，相对路径起始于首个调用文件目录。)

返回值
------
> **类型：**`boolean`  
> **说明：**目录是否存在。

范例
----
>
    <%
    Dim directory
    directory = "./path/to/dir"
    If SE.module("File").dirExists(directory) Then
        Response.Write("目录存在。")
    Else
        Response.Write("目录不存在。")
    End If
    %>