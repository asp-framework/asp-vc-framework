setPanelsStatus
===============
`setPanelsStatus` &mdash; 设置面板状态

说明
----
>     void setPanelsStatus(string panelsStatusString)
> 设置面板状态。

参数
----
> `panelsStatusString`
>> **类型：**`string`  
>> **说明：**面板状态字符串，11个面板按顺序以逗号分割。
>>> 隐藏：0；  
>>> 显示：1。
>>
>> **范例：**`"1,0,0,0,0,0,0,0,0,0,0"`

返回值
------
> 没有返回值。

范例
----
>
    <%
    ' 显示第一个面板
    SE.module("Debugging").setPanelsStatus("1")
>
    ' 显示第一个面板和第三个面板
    SE.module("Debugging").setPanelsStatus("1,0,1")
>
    ' 显示第一个面板
    SE.module("Debugging").setPanelsStatus("1,0,0,0,0,0,0,0,0,0,0")
    %>