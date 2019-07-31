在Excel中，选择菜单“编辑”——“查找(F)…”命令或者按“Ctrl+F”组合键，将弹出如下图01所示的“查找和替换”对话框。在“查找”选项卡中，输入需要查找的内容并设置相关选项后进行查找，Excel会将活动单元格定位在查找到的相应单元格中。如果未发现查找的内容，Excel会弹出“Excel找不到正在搜索的数据”的消息框。
 
图01：“查找”对话框
Excel的这个功能对查找指定的数据非常有用，特别是在含有大量数据的工作表中搜索数据时，更能体现出该功能的快速和便捷。同样，在ExcelVBA中使用与该功能对应的Find方法，提供了一种在单元格区域查找特定数据的简单方式，并且比用传统的循环方法进行查找的速度更快。



1. Find方法的作用
Find方法将在指定的单元格区域中查找包含参数指定数据的单元格，若找到符合条件的数据，则返回包含该数据的单元格；若未发现相匹配的数据，则返回Nothing。该方法返回一个Range对象，在使用该方法时，不影响选定区域或活动单元格。



2. Find方法的语法
[语法]
<单元格区域>.Find (What，[After]，[LookIn]，[LookAt]，[SearchOrder]，[SearchDirection]，[MatchCase]，[MatchByte]，[SearchFormat])
[参数说明]
(1)<单元格区域>，必须指定，返回一个Range对象。
(2)参数What，必需指定。代表所要查找的数据，可以为字符串、整数或者其它任何数据类型的数据。对应于“查找与替换”对话框中，“查找内容”文本框中的内容。
(3)参数After，可选。指定开始查找的位置，即从该位置所在的单元格之后向后或之前向前开始查找(也就是说，开始时不查找该位置所在的单元格，直到Find方法绕回到该单元格时，才对其内容进行查找)。所指定的位置必须是单元格区域中的单个单元格，如果未指定本参数，则将从单元格区域的左上角的单元格之后开始进行查找。
(4)参数LookIn，可选。指定查找的范围类型，可以为以下常量之一：xlValues、xlFormulas或者xlComments，默认值为xlFormulas。对应于“查找与替换”对话框中，“查找范围”下拉框中的选项。
(5)参数LookAt，可选。可以为以下常量之一：XlWhole或者xlPart，用来指定所查找的数据是与单元格内容完全匹配还是部分匹配，默认值为xlPart。对应于“查找与替换”对话框中，“单元格匹配”复选框。
(6)参数SearchOrder，可选。用来确定如何在单元格区域中进行查找，是以行的方式(xlByRows)查找，还是以列的方式(xlByColumns)查找，默认值为xlByRows。对应于“查找与替换”对话框中，“搜索”下拉框中的选项。
(7)参数SearchDirection，可选。用来确定查找的方向，即是向前查找(XlPrevious)还是向后查找(xlNext)，默认的是向后查找。
(8)参数MatchCase，可选。若该参数值为True，则在查找时区分大小写。默认值为False。对应于“查找与替换”对话框中，“区分大小写”复选框。
(9)参数MatchByter，可选。即是否区分全角或半角，在选择或安装了双字节语言时使用。若该参数为True，则双字节字符仅与双字节字符相匹配；若该参数为False，则双字节字符可匹配与其相同的单字节字符。对应于“查找与替换”对话框中，“区分全角/半角”复选框。
(10)参数SearchFormat，可选，指定一个确切类型的查找格式。对应于“查找与替换”对话框中，“格式”按钮。当设置带有相应格式的查找时，该参数值为True。
(11)在每次使用Find方法后，参数LookIn、LookAt、SearchOrder、MatchByte的设置将保存。如果下次使用本方法时，不改变或指定这些参数的值，那么该方法将使用保存的值。
在VBA中设置的这些参数将更改“查找与替换”对话框中的设置；同理，更改“查找与替换”对话框中的设置，也将同时更改已保存的值。也就是说，在编写好一段代码后，若在代码中未指定上述参数，可能在初期运行时能满足要求，但若用户在“查找与替换”对话框中更改了这些参数，它们将同时反映到程序代码中，当再次运行代码时，运行结果可能会产生差异或错误。若要避免这个问题，在每次使用时建议明确的设置这些参数。
3. Find方法使用示例
3.1 本示例在活动工作表中查找what变量所代表的值的单元格，并删除该单元格所在的列。
‘- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Sub Find_Error()
  Dim rng As Range
  Dim what As String
  what = "Error"
  Do
    Set rng = ActiveSheet.UsedRange.Find(what)
    If rng Is Nothing Then
      Exit Do
    Else
       Columns(rng.Column).Delete
    End If
  Loop
End Sub
‘- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
3.2 带格式的查找
本示例在当前工作表单元格中查找字体为"Arial Unicode MS"且颜色为红色的单元格。其中，Application.FindFormat对象允许指定所需要查找的格式，此时Find方法的参数SearchFormat应设置为True。
‘- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Sub FindWithFormat()
  With Application.FindFormat.Font
        .Name = "Arial Unicode MS"
        .ColorIndex = 3
  End With
  Cells.Find(what:="", SearchFormat:=True).Activate
End Sub
‘- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
[小结] 在使用Find方法找到符合条件的数据后，就可以对其进行相应的操作了。您可以：
(1)对该数据所在的单元格进行操作；
(2)对该数据所在单元格的行或列进行操作；
(3)对该数据所在的单元格区域进行操作。



4. 与Find方法相联系的方法
可以使用FindNext方法和FindPrevious方法进行重复查找。在使用这两个方法之前，必须用Find方法指定所需要查找的数据内容。
4.1 FindNext方法
FindNext方法对应于“查找与替换”对话框中的“查找下一个”按钮。可以使用该方法继续执行查找，查找下一个与Find方法中所指定条件的数据相匹配的单元格，返回代表该单元格的Range对象。在使用该方法时，不影响选定区域或活动单元格。
4.1.1 语法
<单元格区域>.FindNext(After)
4.1.2 参数说明
参数After，可选。代表所指定的单元格，将从该单元格之后开始进行查找。开始时不查找该位置所在的单元格，直到FindNext方法绕回到该单元格时，才对其内容进行查找。所指定的位置必须是单元格区域中的单个单元格，如果未指定本参数，则将从单元格区域的左上角的单元格之后开始进行查找。
当查找到指定查找区域的末尾时，本方法将环绕至区域的开始继续查找。发生环绕后，为停止查找，可保存第一次找到的单元格地址，然后测试下一个查找到的单元格地址是否与其相同，作为判断查找退出的条件，以避免出现死循环。当然，如果在查找的过程中，将查找到的单元格数据进行了改变，也可不作此判断，如下例所示。
4.1.3 对VBA帮助系统上的一点疑问探讨
在VBA帮助系统中，介绍Find方法和FindNext方法所使用的示例好像有点问题：当在Excel中运行时，虽然运行结果正确，但是在运行到最后时，会报错误：运行时错误’91’:对象变量或With块变量未设置。究其原因，可能是对象变量c的问题，因为当进行查找并将相应的值全部改变后，最后变量c的值为Nothing。将其稍作改动后，运行通过。
原示例代码如下：(大家也可参见VBA帮助系统Find方法或FindNext方法帮助主题)
本示例在单元格区域A1:A500中查找值为2的单元格，并将这些单元格的值变为5。
With Worksheets(1).Range("a1:a500")
  Set c = .Find(2, lookin:=xlValues)
  If Not c Is Nothing Then
    firstAddress = c.Address
    Do
      c.Value = 5
      Set c = .FindNext(c)
    Loop While Not c Is Nothing And c.Address <> firstAddress
  End If
End With 
经修改后的示例代码如下，即在原代码中加了一句错误处理语句On Error Resume Next，忽略所发生的错误。
Sub test1()
  Dim c As Range, firstAddress As String
  On Error Resume Next
  With Worksheets(1).Range("a1:a15")
    Set c = .Find(2, LookIn:=xlValues)
    If Not c Is Nothing Then
      firstAddress = c.Address
      Do
        c.Value = 5
        Set c = .FindNext(c)
      Loop While Not c Is Nothing And c.Address <> firstAddress
    End If
  End With
End Sub 
或者，将代码作如下修改，即去掉原代码中最后一个判断循环的条件c.Address <> firstAddress，因为本程序的功能是在指定区域查找值为2的单元格并替换为数值5，当程序在指定区域查找不到数值2时就会退出循环，不涉及到重复循环的问题。
Sub test2()
  Dim c As Range, firstAddress As String
  With Worksheets(1).Range("a1:a15")
    Set c = .Find(2, LookIn:=xlValues)
    If Not c Is Nothing Then
      firstAddress = c.Address
      Do
        c.Value = 5
        Set c = .FindNext(c)
      Loop While Not c Is Nothing
    End If
  End With
End Sub 
您也可以试试该程序，看看我的理解是否正确，或者还有什么其它的解决办法。
4.2 FindPrevious方法
可以使用该方法继续执行Find方法所进行的查找，查找前一个与Find方法中所指定条件的数据相匹配的单元格，返回代表该单元格的Range对象。在使用该方法时，不影响选定区域或活动单元格。
4.2.1 语法
<单元格区域>.FindPrevious(After)
4.2.2 参数说明
参数After，可选。代表所指定的单元格，将从该单元格之前开始进行查找。开始时不查找该位置所在的单元格，直到FindPrevious方法绕回到该单元格时，才对其内容进行查找。所指定的位置必须是单元格区域中的单个单元格，如果未指定本参数，则将从单元格区域的左上角的单元格之前开始进行查找。
当查找到指定查找区域的起始位置时，本方法将环绕至区域的末尾继续查找。发生环绕后，为停止查找，可保存第一次找到的单元格地址，然后测试下一个查找到的单元格地址是否与其相同，作为判断查找退出的条件，以避免出现死循环。
4.2.3 示例
在工作表中输入如下图02所示的数据，至少保证在A列中有两个单元格输入了数据“excelhome”。
 图02：测试的数据
在VBE编辑器中输入下面的代码测试Find方法、FindNext方法、FindPrevious方法，体验各个方法所查找到的单元格位置。
‘- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Sub testFind()
  Dim findValue As Range
  Set findValue = Worksheets("Sheet1").Columns("A").Find(what:="excelhome")
  MsgBox "第一个数据发现在单元格:" & findValue.Address
  Set findValue = Worksheets("Sheet1").Columns("A").FindNext(After:=findValue)
  MsgBox "下一个数据发现在单元格:" & findValue.Address
  Set findValue = Worksheets("Sheet1").Columns("A").FindPrevious(After:=findValue)
  MsgBox "前一个数据发现在单元格" & findValue.Address
End Sub

