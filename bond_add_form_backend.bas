Option Explicit

Private Sub addBond_Click()
With Worksheets("bond_portfolio_data")
Dim cp As Double, md As Date, ppy As Integer, fv As Currency
Dim cr As Double, br As String, bt As String, dr As Double
Range("A2").EntireRow.Insert
Range("A2").EntireRow.ClearFormats
Range("A2").value = CCur(fvEntry)
Range("A2").NumberFormat _
= "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("B2").value = CDate(mdEntry)
Range("C2").value = CDec(crEntry)
Range("D2").value = CInt(ppyEntry)
Range("E2").value = CStr(brEntry)
Range("F2").value = CStr(btEntry)
Range("G2").value = CDec(drEntry)
End With
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
