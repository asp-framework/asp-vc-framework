getIncludeResult
================
`getIncludeResult` &mdash; 包含文件并获取执行后的内容(不输出内容)

说明
----
>     string getIncludeResult(ByVal filePath)
> 包含文件并获取执行后的内容。(不输出内容。)

参数
----
> `filePath`
>> **类型：**`string`  
>> **说明：**文件路径。(路径格式为相对路径，相对路径起始于首个调用文件目录。)  
>> **范例：**`"dir/filePath"`

返回值
------
> **类型：**`string`  
> **说明：**执行后的字符串内容。

范例
----
>
    目录文件
>
    Dir/
        file.asp
    index.asp
>>
>
    <%
    '''
     ' Dir/file.asp 文件
     ''
    Response.Write("成功执行。")
    %>
>>
>
    <%
    '''
     ' index.asp 文件
     ''
    Dim resultString
    resultString = SE.getIncludeResult("Dir/file.asp")
    Response.Write(resultString)
    %>
> 运行`index.asp`将得到以下内容：  
>
    成功执行。