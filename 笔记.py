Set Values = ActiveSheet.Range("D2:D102")
ActiveSheet.Range("D" & num)
Range('A1:B5').ReplaceWhat:='A', Replacement:='MM', MatchCase:=True
      
      
当加上On Error Resume Next语句后，如果后面的程序出现"运行时错误"时，会继续运行，不中断。
当加上On Error Goto 0语句后，如果后面的程序出现"运行时错误"时，会显示"出错信息"并停止程序的执行。
