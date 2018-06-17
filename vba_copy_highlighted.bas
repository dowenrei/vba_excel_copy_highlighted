Sub CopyHighlightedCell()
'Declare Variable
Dim LR As Long, j As Long, colour As Long

Dim c As Range
Dim rReply As Range

'Declare New Worksheet
Worksheets.Add After:=ActiveSheet
ActiveSheet.Name = "Sheet2"
Worksheets("Sheet1").Activate

'Get Highlighted Cell Colour from User Input
Set rReply = Application.InputBox(Prompt:="Selct a single cell that has the background color you wish to copy", Type:=8)
colour = rReply.Interior.ColorIndex

'Copy Highlighted Cell to another sheet
j = 1
LR = Range("A" & Rows.Count).End(xlUp).Row
For Each c In ActiveSheet.UsedRange
      If c.Interior.ColorIndex = colour Then
            c.Copy Destination:=Worksheets("Sheet2").Range("A" & j)
            j = j + 1
        End If
Next c
End Sub
