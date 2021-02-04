Attribute VB_Name = "Korpus"
Public Sub IngredientsList_copy()

Dim i As Long
Dim cr As Long
Dim firstRow As Long
Dim lastRow As Long
Dim x As Byte
Dim r As Integer
Dim cf As Integer
Dim j As Double
Dim y As Integer
Dim pot As Integer
Dim z As Integer
Dim f As Integer


r = 5
x = 0


cr = 0

cr = List12.Range("F2").Value2
List12.Range("E2").Value = cr

cf = List12.Range("M4")
pot = List12.Range("P4")

For i = 1 To 300
 Do While List25.Cells(i, 1).Value = cr
  If List25.Cells(i, 1).Value = cr Then
  x = x + 1
  List12.Range("M2").Value = x
  List25.Rows(i).EntireRow.Copy List12.Rows(r + x)
  
  'Formáty
  
  
  List12.Range("C6:I32").Select
  With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
  End With
  
    List12.Range("C6:I32").Select
  With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
  End With
    
    List12.Range("D6:I28").Select
  With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
  End With
        
   'Konec formátù
        
  lastRow = r + x
  Exit Do
  End If
 Loop
Next i
 
If lastRow <= 5 Then
 lastRow = 6
End If
 
 
lastRow = lastRow + 1
 
For j = lastRow To 21
 List12.Rows(j).ClearContents
 List12.Rows(j).Borders(xlEdgeBottom).LineStyle = xlNone
 List12.Range("M3").Value = lastRow
Next j

 List12.Rows(lastRow).Select
 With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThick
 End With
 
 
List12.Range("M5") = cf
y = 33
 
If cf > 0 Then
    For j = 1 To 400
    Do While List20.Cells(j, 1).Value = cf
        If List20.Cells(j, 1).Value = cf Then
        y = y + 1
        List12.Range("M5").Value = cf
        List20.Rows(j).EntireRow.Copy List12.Rows(y)
        Exit Do
        End If
        Loop
    Next j
End If

f = 33

If pot > 0 Then
    For z = 1 To 400
    Do While List20.Cells(z, 1).Value = pot
        If List20.Cells(z, 1).Value = pot Then
        f = f + 1
        List12.Range("P5").Value = pot
        List20.Rows(z).EntireRow.Copy List12.Rows(f)
        Exit Do
        End If
        Loop
    Next z
End If

If cf = 0 And pot = 0 Then
   For l = 33 To 40
   List12.Rows(l).Clear

   Next l
    
   Range("C34:I40").Select
   With Selection.Borders
        .LineStyle = xlNone
   End With
   
   Selection.Borders(xlDiagonalDown).LineStyle = xlNone
   Selection.Borders(xlDiagonalUp).LineStyle = xlNone
     
   
End If
        
End Sub

'@ Radek Dockalik 2020


