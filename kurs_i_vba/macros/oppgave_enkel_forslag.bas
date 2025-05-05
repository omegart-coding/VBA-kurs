Attribute VB_Name = "kurs_enkel"
Option Explicit
Function MostOccur(totalArea As Range) As String

Dim row As Integer
Dim Occurances(1 To 3) As Integer

Occurances(1) = 0 'R
Occurances(2) = 0 'G
Occurances(3) = 0 'B
'Declares occurances as an array

For row = 1 To totalArea.Rows.Count
If Not IsEmpty(totalArea.Cells(row)) Then
Dim curValue
curValue = totalArea.Cells(row).Value

If curValue = "R" Then
Occurances(1) = Occurances(1) + 1
ElseIf curValue = "G" Then
Occurances(2) = Occurances(2) + 1
ElseIf curValue = "B" Then
Occurances(3) = Occurances(3) + 1
End If
End If

Next row

If Occurances(1) > Occurances(2) And Occurances(1) > Occurances(3) Then
MostOccur = "R"
ElseIf Occurances(2) > Occurances(1) And Occurances(2) > Occurances(3) Then
MostOccur = "G"
ElseIf Occurances(3) > Occurances(1) And Occurances(3) > Occurances(2) Then
MostOccur = "B"
End If

End Function
