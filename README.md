# PingToExcel
# Ping computers on a network by name, and the export the results to excel.

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

objExcel.Workbooks.Add

intRow = 2

 

objExcel.Cells(1, 1).Value = "Machine Name"

objExcel.Cells(1, 2).Value = "Results"

 

Set Fso = CreateObject("Scripting.FileSystemObject")

Set InputFile = fso.OpenTextFile("C:\Users\trp\Desktop\computernames.txt")

 

Do While Not (InputFile.atEndOfStream)

HostName = InputFile.ReadLine

 

Set WshShell = WScript.CreateObject("WScript.Shell")

Ping = WshShell.Run("ping -n 1 " & HostName, 0, True)

 

objExcel.Cells(intRow, 1).Value = UCase(HostName)

 

Select Case Ping

Case 0 objExcel.Cells(intRow, 2).Value = "On Line"

Case 1 objExcel.Cells(intRow, 2).Value = "Off Line"

End Select

 

If objExcel.Cells(intRow, 2).Value = "Off Line" Then

objExcel.Cells(intRow, 2).Interior.ColorIndex = 3

Else

objExcel.Cells(intRow, 2).Interior.ColorIndex = 4

End If

 

intRow = intRow + 1

Loop

 

objExcel.Range("A1:B1").Select

objExcel.Selection.Interior.ColorIndex = 19

objExcel.Selection.Font.ColorIndex = 11

objExcel.Selection.Font.Bold = True

objExcel.Cells.EntireColumn.AutoFit

 

MsgBox "Done"

 

Entire Row:

 

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

objExcel.Workbooks.Add

intRow = 2

 

objExcel.Cells(1, 1).Value = "Machine Name"

objExcel.Cells(1, 2).Value = "Results"

 

Set Fso = CreateObject("Scripting.FileSystemObject")

Set InputFile = fso.OpenTextFile("MachineList.Txt")

 

Do While Not (InputFile.atEndOfStream)

HostName = InputFile.ReadLine

 

Set WshShell = WScript.CreateObject("WScript.Shell")

Ping = WshShell.Run("ping -n 1 " & HostName, 0, True)

 

objExcel.Cells(intRow, 1).Value = UCase(HostName)

 

Select Case Ping

Case 0 objExcel.Cells(intRow, 2).Value = "On Line"

Case 1 objExcel.Cells(intRow, 2).Value = "Off Line"

End Select

 

If objExcel.Cells(intRow, 2).Value = "Off Line" Then

objExcel.Cells(intRow, 2).EntireRow.Interior.ColorIndex = 3

Else

objExcel.Cells(intRow, 2).EntireRow.Interior.ColorIndex = 4

End If

 

intRow = intRow + 1

Loop

 

objExcel.Range("A1:B1").Select

objExcel.Selection.Interior.ColorIndex = 19

objExcel.Selection.Font.ColorIndex = 11

objExcel.Selection.Font.Bold = True

objExcel.Cells.EntireColumn.AutoFit

 

MsgBox "Fin"

 

Font Color:

 

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True

objExcel.Workbooks.Add

intRow = 2

 

objExcel.Cells(1, 1).Value = "Machine Name"

objExcel.Cells(1, 2).Value = "Results"

 

Set Fso = CreateObject("Scripting.FileSystemObject")

Set InputFile = fso.OpenTextFile("MachineList.Txt")

 

Do While Not (InputFile.atEndOfStream)

HostName = InputFile.ReadLine

 

Set WshShell = WScript.CreateObject("WScript.Shell")

Ping = WshShell.Run("ping -n 1 " & HostName, 0, True)

 

objExcel.Cells(intRow, 1).Value = UCase(HostName)

 

Select Case Ping

Case 0 objExcel.Cells(intRow, 2).Value = "On Line"

Case 1 objExcel.Cells(intRow, 2).Value = "Off Line"

End Select

 

If objExcel.Cells(intRow, 2).Value = "Off Line" Then

objExcel.Cells(intRow, 2).Font.ColorIndex = 3

Else

objExcel.Cells(intRow, 2).Font.ColorIndex = 4

End If

 

intRow = intRow + 1

Loop

 

objExcel.Range("A1:B1").Select

objExcel.Selection.Interior.ColorIndex = 19

objExcel.Selection.Font.ColorIndex = 11

objExcel.Selection.Font.Bold = True

objExcel.Cells.EntireColumn.AutoFit

 

MsgBox "Finished"


