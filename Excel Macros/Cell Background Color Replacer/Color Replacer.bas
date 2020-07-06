Attribute VB_Name = "ColorReplacer"
'Created by Rimoun George
Public Sub ColorReplacer()
Dim NewColor As Range
Dim OldColor As Range
Dim RangeReplace As Range

On Error GoTo ErrHandler
' Select One Cell of Each Color
Set NewColor = Application.InputBox("New Color", "Select The Cell that has the color you want to apply", Type:=8)
Set OldColor = Application.InputBox("Old Color", "Select The Cell that has the color you want to replace", Type:=8)
' Select once Cell or more to apply the replacement
Set RangeReplace = Application.InputBox("Select the range of Cells", "Select The Cells that you want to apply the new color on them", Type:=8)

'Checks the number of cells of each range before applying the action
If NewColor.Cells.Count = 1 And OldColor.Cells.Count = 1 And RangeReplace.Cells.Count > 0 Then

NewClr = getRGB2(NewColor)
OldClr = getRGB2(OldColor)
NewColorVars = Split(NewClr, ",")

For Each cell In RangeReplace
If getRGB2(cell) = OldClr Then cell.Interior.Color = RGB(NewColorVars(0), NewColorVars(1), NewColorVars(2))
Next

Else
GoTo ErrHandler
End If

Exit Sub

ErrHandler:
MsgBox "Wrong Ranges were Selected!" & vbCrLf & "============================" & vbCrLf & "New Color: 1 Cell only" & vbCrLf & "Old Color: 1 Cell only" & vbCrLf & "Replace Range: 1 cell or more", vbCritical, "Error"
End Sub



Private Function getRGB2(rcell) As String
    Dim C As Long
    Dim R As Long
    Dim G As Long
    Dim B As Long

    C = rcell.Interior.Color
    R = C Mod 256
    G = C \ 256 Mod 256
    B = C \ 65536 Mod 256
    getRGB2 = R & "," & G & "," & B
End Function
