Attribute VB_Name = "Module1"
Public RowIndex, ColIndex As Integer

Sub tech()
Call ClearContents

Dim i, j As Integer
i = 5
j = 3
Dim a As Integer
Dim tech As String
For a = 3 To 18
    If Sheet1.Range("AA" & a) <> "" Then
        tech = Sheet1.Range("AA" & a).Text
        Sheet12.Range("b" & i) = tech
        Sheet12.Cells(4, j) = tech
        i = i + 1
        j = j + 1
    Else
        Exit For
        
    End If
Next
Sheet12.Activate
End Sub

Sub ClearContents()
Dim row, col As Integer
For row = 5 To 14
    Sheet12.Range("b" & row).ClearContents
    Sheet12.Range("D5:L5").ClearContents
    Sheet12.Range("E6:L6").ClearContents
    Sheet12.Range("F7:L7").ClearContents
    Sheet12.Range("G8:L8").ClearContents
    Sheet12.Range("H9:L9").ClearContents
    Sheet12.Range("I10:L10").ClearContents
    Sheet12.Range("J11:L11").ClearContents
    Sheet12.Range("K12:L12").ClearContents
    Sheet12.Range("L13").ClearContents
Next
For col = 3 To 12
    Sheet12.Cells(4, col).ClearContents
    
Next
End Sub



Sub test()
Dim percentage As String
End Sub
Sub tech2()
Call ClearContents2

Dim i, j As Integer
i = 5
j = 3
Dim a As Integer
Dim tech2 As String
For a = 3 To 18
    If Sheet1.Range("AA" & a) <> "" Then
        tech2 = Sheet1.Range("AA" & a).Text
        Sheet14.Range("b" & i) = tech2
        Sheet14.Cells(4, j) = tech2
        i = i + 1
        j = j + 1
    Else
        Exit For
        
    End If
Next
Sheet14.Activate
End Sub

Sub ClearContents2()
Dim row, col As Integer
For row = 5 To 14
    Sheet14.Range("b" & row).ClearContents
    Sheet14.Range("D5:L5").ClearContents
    Sheet14.Range("E6:L6").ClearContents
    Sheet14.Range("F7:L7").ClearContents
    Sheet14.Range("G8:L8").ClearContents
    Sheet14.Range("H9:L9").ClearContents
    Sheet14.Range("I10:L10").ClearContents
    Sheet14.Range("J11:L11").ClearContents
    Sheet14.Range("K12:L12").ClearContents
    Sheet14.Range("L13").ClearContents
Next
For col = 3 To 12
    Sheet14.Cells(4, col).ClearContents
Next
End Sub

Sub tech3()
Call ClearContents3

Dim i, j As Integer
i = 5
j = 3
Dim a As Integer
Dim tech3 As String
For a = 3 To 18
    If Sheet1.Range("AA" & a) <> "" Then
        tech3 = Sheet1.Range("AA" & a).Text
        Sheet16.Range("b" & i) = tech3
        Sheet16.Cells(4, j) = tech3
        i = i + 1
        j = j + 1
    Else
        Exit For
        
    End If
Next
Sheet16.Activate
End Sub

Sub ClearContents3()
Dim row, col As Integer
For row = 5 To 14
    Sheet16.Range("b" & row).ClearContents
    Sheet16.Range("D5:L5").ClearContents
    Sheet16.Range("E6:L6").ClearContents
    Sheet16.Range("F7:L7").ClearContents
    Sheet16.Range("G8:L8").ClearContents
    Sheet16.Range("H9:L9").ClearContents
    Sheet16.Range("I10:L10").ClearContents
    Sheet16.Range("J11:L11").ClearContents
    Sheet16.Range("K12:L12").ClearContents
    Sheet16.Range("L13").ClearContents
Next
For col = 3 To 12
    Sheet16.Cells(4, col).ClearContents
Next
End Sub


Sub ROI()
Dim capa(1 To 5) As String
capa(1) = "People"
capa(2) = "Facilities"
capa(3) = "Spares"
capa(4) = "Test equipment"
capa(5) = "Information"

Dim a, i, counter, j As Integer
i = 2
counter = 0
j = 2
For a = 3 To 18
    If Sheet1.Range("AA" & a) <> "" Then
        Sheet15.Range("a" & i) = Sheet1.Range("AA" & a).Text
        i = i + 5
        counter = counter + 1
    Else
        Exit For
    End If
        
Next
For j = 2 To counter * 5 + 1
    Sheet15.Range("b" & j) = capa(((j - 2) Mod 5) + 1)
Next
Dim rng As Range

For Each rng In Sheet15.Range("A1:A" & counter * 5)
   If rng = Sheet3.Range("B2").Text Then
      rng.Offset(0, 2) = Sheet3.Range("L2")
      rng.Offset(1, 2) = Sheet3.Range("L3")
      rng.Offset(2, 2) = Sheet3.Range("L4")
      rng.Offset(3, 2) = Sheet3.Range("L5")
      rng.Offset(4, 2) = Sheet3.Range("L6")
      rng.Offset(0, 3) = Sheet3.Range("N2")
      Exit For
   End If


Next


End Sub
Sub ClearContents4()

Sheet15.Range("C2:D200").ClearContents
Sheet15.Range("A2:B200").ClearContents

End Sub
Sub ClearContents5()

Sheet18.Range("A2:C200").ClearContents

End Sub



Sub ROI2()
Dim yea(1 To 5) As String
yea(1) = "Year1"
yea(2) = "Year2"
yea(3) = "Year3"
yea(4) = "Year4"
yea(5) = "Year5"

Dim b, i, counter, j As Integer
i = 2
counter = 0
j = 2
For b = 3 To 18
    If Sheet1.Range("AA" & b) <> "" Then
        Sheet18.Range("a" & i) = Sheet1.Range("AA" & b).Text
        i = i + 5
        counter = counter + 1
    Else
        Exit For
    End If
        
Next
For j = 2 To counter * 5 + 1
    Sheet18.Range("b" & j) = yea(((j - 2) Mod 5) + 1)
Next
Dim rng As Range

For Each rng In Sheet18.Range("A2:A" & counter * 5)
   If rng = Sheet3.Range("A20").Text Then
      rng.Offset(0, 2) = Sheet3.Range("D25").Text
      rng.Offset(1, 2) = Sheet3.Range("E25").Text
      rng.Offset(2, 2) = Sheet3.Range("F25").Text
      rng.Offset(3, 2) = Sheet3.Range("G25").Text
      rng.Offset(4, 2) = Sheet3.Range("K25").Text
    
      Exit For
   End If


Next


End Sub

Sub highlight()
If ColIndex = 0 Then GoTo Line
Sheet7.Cells(RowIndex, ColIndex).Select
Call ClearHighlight
Line:
Dim str1, str2 As String
str1 = Sheet13.Range("AC7").Text
str2 = Sheet13.Range("AD7").Text

For RowIndex = 4 To 34
    If Sheet7.Range("D" & RowIndex) = str1 Then
        Exit For
    End If
    
Next
For ColIndex = 6 To 9
    If Sheet7.Cells(2, ColIndex) = str2 Then
        Exit For
    End If
    
Next
If ColIndex = 6 Then
    ColIndex = ColIndex - 1
End If

Sheet7.Cells(RowIndex, ColIndex).Interior.Color = 255
Sheet7.Cells(RowIndex, ColIndex).Select
'MsgBox (RowIndex)
'MsgBox (ColIndex)

End Sub

Sub CloseClear()

Dim rng As Range
For Each rng In Sheet7.Range("E4:I36")
    If rng.Interior.Color = 255 Then
        rng.Interior.Color = rng.Offset(2, 0).Interior.Color
    End If
Next

End Sub

