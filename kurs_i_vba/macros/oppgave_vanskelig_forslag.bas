Attribute VB_Name = "kurs_test"
Option Explicit
Function readCells(totalArea As Range, desVal As String) As Integer

Dim col As Integer, row As Integer, i As Integer, total As Integer

Dim Occurances(1 To 3) As Integer

Occurances(1) = 0 'R
Occurances(2) = 0 'G
Occurances(3) = 0 'B
'Declares occurances as an array

total = 0

For col = 1 To totalArea.Columns.Count

For row = 1 To totalArea.Rows.Count
If Not IsEmpty(totalArea.Cells(row, col)) Then
Dim curValue
curValue = totalArea.Cells(row, col).Value
If curValue = "R" Then
Occurances(1) = Occurances(1) + 1

ElseIf curValue = "G" Then
Occurances(2) = Occurances(2) + 1

ElseIf curValue = "B" Then
Occurances(3) = Occurances(3) + 1
End If
'If cell value exist in array, increases the counting of the array for the corresponding value
End If

Next row

Dim WanVal As Integer, OtherVal1 As Integer, OtherVal2 As Integer

If desVal = "R" Then
WanVal = 1
OtherVal1 = 2
OtherVal2 = 3

ElseIf desVal = "G" Then
WanVal = 2
OtherVal1 = 1
OtherVal2 = 3

ElseIf desVal = "B" Then
WanVal = 3
OtherVal1 = 1
OtherVal2 = 2
End If
'Declares which values is unwanted

If Occurances(WanVal) > Occurances(OtherVal1) And Occurances(WanVal) > Occurances(OtherVal2) Then
total = total + 1
'If the desired value is the most frequent, increases the total count
End If

Occurances(1) = 0
Occurances(2) = 0
Occurances(3) = 0
'Resets the array

Next col

readCells = total

End Function

