VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBox1_Change()

myVal = Me.ComboBox1.Value

lr = ThisWorkbook.Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row

Me.ComboBox2.Clear
Me.ComboBox2.Text = Empty

For x = 2 To lr
    If myVal = ThisWorkbook.Sheets("Sheet1").Cells(x, 2) Then
    
        Me.ComboBox2.AddItem ThisWorkbook.Sheets("Sheet1").Cells(x, "G")
    
    End If
Next x

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Call FillCombos
Sheet1.ComboBox2.ListFillRange = "IDno"
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
