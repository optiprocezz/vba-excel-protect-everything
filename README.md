# vba-excel-protect-everything

```vb
Function RandomString(Length As Integer)
    Dim characterBank As Variant
    Dim x As Long
    Dim str As String

    characterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
        "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
        "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "!", "@", _
        "#", "$", "%", "^", "&", "*", "A", "B", "C", "D", "E", "F", "G", "H", _
        "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
        "W", "X", "Y", "Z")

    'Randomly Select Characters One-by-One
    For x = 1 To Length
        Randomize
        str = str & characterBank(Int((UBound(characterBank) - LBound(characterBank) + 1) * Rnd + LBound(characterBank)))
    Next x

    'Output Randomly Generated String
    RandomString = str

End Function


Sub lockDocument()
    ' Create a variable to hold worksheets
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the active workbook
    For Each ws In ActiveWorkbook.Worksheets
        
        ' Disable selection of cells
        ws.EnableSelection = xlNoSelection
    
        ' Protect the worksheet
        ws.Protect password:=RandomString(20)
    
    Next ws
    
    ' Protect the workbook
    ActiveWorkbook.Protect password:=RandomString(20)

End Sub
```
