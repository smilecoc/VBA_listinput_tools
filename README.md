# VBA_listinput_tools
##公众号：Romi的杂货铺，如有疑问与交流欢迎关注公众号！

Excel利用VBA实现下拉列表，同时支持输入时动态查询，根据输入的不同实现动态的查询

先看一下实验效果：

当点击website这一列时会出现所有的网站列表，双击可点击选择数值填入

![image](https://upload-images.jianshu.io/upload_images/16636256-90568c599ef8bba5?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

输入关键字时会只出现包含关键字的结果

![image](https://upload-images.jianshu.io/upload_images/16636256-5f2be809c6a83021?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

在C，D两列选择单元格后会出现仅在此网站下的数据如果网站为空则会自动向上寻找，同时也支持自定义的搜索

![image](https://upload-images.jianshu.io/upload_images/16636256-41f8ce51eaf63dee?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

![image](https://upload-images.jianshu.io/upload_images/16636256-1c85dc19a4b1fdfe?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

接下来为主要的实现方法：

第一部分为工作表选取改变事件，实现的是当有单元格被选定时会自动出现下拉菜单和输入框。首先需要在sheet中创建一个listbox和textbox.在开发工具-插入-下拉框/文本框注意要选activex控件，不能选择上面的控件
![image.png](https://upload-images.jianshu.io/upload_images/16636256-a05955d95a177b9a.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

具体代码及注释如下：

```
'工作表选取改变事件
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim i, x, rownu As Variant
    Dim d As Object
    Dim arr, arr_key, arr1, yun, arr_po
    Dim website_name As String
    
    Set d = CreateObject("scripting.dictionary")
    Me.ListBox1.Clear
    'target为选取的单元格对象
    tacolumn = Target.Column
    tarow = Target.Row
  
    '添加website部分
    '选择触发的区域，使用Target.Cells.CountLarge是为了保证选择的是一个单元格而不是一片区域，同时区域过大不会报错
        If Target.Column = 1 And Target.Row > 10 And Target.Cells.CountLarge = 1 Then
            With Me.TextBox1'textbox的大小，位置，和显示
                .Visible = True
                .Top = Target.Top
                .Left = Target.Left
                .Width = Target.Width
                .Height = Target.Height
                .Activate
            End With
            With Me.ListBox1'listbox的大小，位置，和显示
                .Visible = True
                .Top = Target.Top
                .Left = Target.Left + Target.Width
                .Width = 400
                .Height = 300
                '将需要写入的数据装入数组
                arr = Sheets("format_information").Range("a2:a" & Sheets("format_information").Cells(Rows.Count, 1).End(xlUp).Row)
                For x = 1 To UBound(arr)
                d(arr(x, 1)) = ""
                Next
                '将值写入到listbox中
                .List = d.keys()
                
            End With
    
     'position和fomat部分.逻辑与上述代码一致
        ElseIf (Target.Column = 3 Or Target.Column = 4) And Target.Row > 10 And Target.Cells.CountLarge = 1 Then
            website_name = Cells(Target.Row, 1).Value
            rownu = Target.Row - 1
            Do Until website_name <> ""
                website_name = Cells(rownu, 1).Value
                rownu = rownu - 1
            Loop
            
            With Me.TextBox1
                .Visible = True
                .Top = Target.Top
                .Left = Target.Left
                .Width = Target.Width
                .Height = Target.Height
                .Activate
            End With
            With Me.ListBox1
                .Visible = True
                .Top = Target.Top
                .Left = Target.Left + Target.Width
                .Width = 400
                .Height = 300
                yun = SQLtoArr("Select position_channel,Format FROM [format_information$] where Website like '%" & website_name & "%'")
                arr_po = Sheets("format_information").Range("AA1:AA" & Sheets("format_information").Cells(Rows.Count, 27).End(xlUp).Row)
                arr1 = Sheets("format_information").Range("AB1:AB" & Sheets("format_information").Cells(Rows.Count, 28).End(xlUp).Row)
                For x = 1 To UBound(arr_po)
                d(arr_po(x, 1) & "■" & arr1(x, 1)) = ""
                Next
                .List = d.keys()
            
            End With
                      
        
        
        Else
            Me.ListBox1.Clear
            Me.TextBox1 = ""
            Me.ListBox1.Visible = False
            Me.TextBox1.Visible = False
        End If
    
End Sub
```

```
'利用SQL函数进行筛选和取值的函数

Function SQLtoArr(strSQL)

 Dim Conn As Object, Rst As Object
 Dim strConn As String
 Dim i As Integer, PathStr As String
 Set Conn = CreateObject("ADODB.Connection")
 Set Rst = CreateObject("ADODB.Recordset")
 PathStr = ThisWorkbook.FullName '设置工作簿的完整路径和名称
 Select Case Application.Version * 1 '设置连接字符串,根据版本创建连接
 Case Is <= 11
    strConn = "Provider=Microsoft.Jet.Oledb.4.0;Extended Properties=excel 8.0;Data source=" & PathStr
 Case Is >= 12
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & PathStr & ";Extended Properties=""Excel 12.0;HDR=YES"";"""
 End Select
 
Conn.Open strConn '打开数据库链接
Set Rst = Conn.Execute(strSQL) '执行查询，并将结果输出到记录集对象
Sheets("format_information").Columns("AA:AB").Clear
Sheets("format_information").Range("AA2").CopyFromRecordset Rst '#####################在这里改输出的位置与单元格
Rst.Close  '关闭数据库连接
Conn.Close
'Set Conn = Nothing
'Set Rst = Nothing


End Function
```

第二部分为键入字符后执行搜索的功能

```

'textbox键盘抬起事件：即输入了文字
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim i, x As Integer
    Dim Language As Boolean, arr1 As Variant
    Dim myStr As String, str_B As String
    Dim d As Object
    Dim arr, arr_key
    
    Set d = CreateObject("scripting.dictionary")
    Me.ListBox1.Clear
    myStr = Me.TextBox1.Value
    With Me.ListBox1
                .Width = 400
                .Height = 300
    End With
    If tacolumn = 1 And tarow > 10 Then
    With Sheets("format_information")
           
                arr1 = .Range("a2:a" & .Range("a65535").End(xlUp).Row)
                For i = 1 To .Range("a65535").End(xlUp).Row - 1
                '利用instr遍历找到包含输入文字的部分,并 赋值到字典里避免重复
                   If InStr(1, arr1(i, 1), myStr, 1) Then
                       d(arr1(i, 1)) = ""
                   End If
                Next i
                Me.ListBox1.List = d.keys()'listbox赋值
            
    End With
    ElseIf (tacolumn = 3 Or tacolumn = 4) And tarow > 10 Then
    With Sheets("format_information")
           
                arr = .Range("c2:c" & .Range("c65535").End(xlUp).Row)
                arr1 = .Range("d2:d" & .Range("d65535").End(xlUp).Row)
                For i = 1 To .Range("c65535").End(xlUp).Row - 1
                   If InStr(1, arr(i, 1), myStr, 1) Or InStr(1, arr1(i, 1), myStr, 1) Then
                       d(arr(i, 1) & "■" & arr1(i, 1)) = ""
                   End If
                Next i
                
                Me.ListBox1.List = d.keys()
                
    End With
    End If
End Sub
```

第三部分为双击选取值的部分

```

'listbox双击事件
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim arr
    
    If (tacolumn = 1 Or tacolumn = 2) And tarow > 10 Then
    '将listbox值赋予当前单元格
        ActiveCell.Value = Me.ListBox1.Value
        Me.ListBox1.Clear
        Me.TextBox1 = ""'清空listbox与textbox
        Me.ListBox1.Visible = False'y隐藏textbox和listbox
        Me.TextBox1.Visible = False
     ElseIf (tacolumn = 3 Or tacolumn = 4) And tarow > 10 Then
        arr = Split(Me.ListBox1.Value, "■")
        ActiveCell.Value = arr(0)
        ActiveCell.Offset(0, 1).Value = arr(1)
        Me.ListBox1.Clear
        Me.TextBox1 = ""
        Me.ListBox1.Visible = False
        Me.TextBox1.Visible = False
    End If
End Sub
```


