Attribute VB_Name = "Módulo1"
Sub impressao()

Worksheets("Nao_Conformidades").Activate

lastRow = 2
rTotal = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
For x = rTotal To 2 Step -1
    If ActiveSheet.Cells(x, "B").Value <> "" Then
        lastRow = x
        x = 2
    End If
Next x

ActiveSheet.PageSetup.PrintArea = "$B$1:$I$" & lastRow
    
End Sub
