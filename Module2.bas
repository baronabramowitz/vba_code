Option Explicit

Sub BondMetrics()

Dim cp As Double, md As Date, ppy As Integer, cpd As Date, cpdMod As Date
Dim cr As Double, br As String, bt As String, wd As Integer, dr As Double
Dim fv As Currency, td As Date, cpds() As Variant, i As Integer, integ As Integer
Dim dtp As Long, dtps() As Variant, pvcps() As Variant, pvfcfs As Double
Dim durcalc() As Variant, durcalcs As Double
Dim convcalc() As Variant, convcalcs As Double

td = Now
'fv = Sheet1.Range("$A2").value
'md = Sheet1.Range("$B2").value
'cr = Sheet1.Range("$C2").value
'ppy = Sheet1.Range("$D2").value
'br = Sheet1.Range("$E2").value
'bt = Sheet1.Range("$F2").value
'dr = Sheet1.Range("$G2").value
fv = ActiveCell.value
md = ActiveCell.Offset(0, 1).value
cr = ActiveCell.Offset(0, 2).value
ppy = ActiveCell.Offset(0, 3).value
br = ActiveCell.Offset(0, 4).value
bt = ActiveCell.Offset(0, 5).value
dr = ActiveCell.Offset(0, 6).value
If ppy = 0 Then
    cp = 0
Else:
    cp = (fv * cr) / ppy
End If
If Weekday(md, vbMonday) = 7 Then
    md = md + 1
ElseIf Weekday(md, vbMonday) = 6 Then
    md = md + 2
End If
If ppy = 0 Then
    pvfcfs = fv / (1 + dr / 365) ^ (md - DateValue(CStr(Now())))
    durcalcs = (md - td) / 365
    convcalcs = (((md - td) / 365) ^ 2 + (md - td) / 365) / (1 + dr) ^ 2
Else:
    cpd = DateAdd("M", -12 / ppy, md)
    If Weekday(cpd, vbMonday) = 7 Then
        cpd = cpd + 1
    ElseIf Weekday(cpd, vbMonday) = 6 Then
        cpd = cpd + 2
    End If
    ReDim Preserve cpds(0 To 1)
    cpds(0) = md
    cpds(1) = cpd
    integ = 2
    
    Do
        ReDim Preserve cpds(0 To integ)
        cpd = DateAdd("M", -12 / ppy, cpd)
            If Weekday(cpd, vbMonday) = 7 Then
                cpdMod = cpd + 1
                cpds(integ) = cpdMod
            ElseIf Weekday(cpd, vbMonday) = 6 Then
                cpdMod = cpd + 2
                cpds(integ) = cpdMod
            Else
                cpds(integ) = cpd
            End If
        integ = integ + 1
    Loop Until cpd < td
    ReDim Preserve cpds(0 To UBound(cpds) - 1)
    For i = 0 To UBound(cpds)
        ReDim Preserve dtps(0 To i)
        dtps(i) = cpds(i) - DateValue(CStr(Now()))
    Next i
    For i = 0 To UBound(dtps)
        ReDim Preserve pvcps(0 To i)
        pvcps(i) = cp / (1 + dr / 365) ^ dtps(i)
    Next i
    For i = 0 To UBound(dtps)
    ReDim Preserve durcalc(0 To i)
    If i = 0 Then
        durcalc(i) = (dtps(i) / 365) * (pvcps(i) + fv / (1 + dr / 365) ^ dtps(0))
    Else
        durcalc(i) = (dtps(i) / 365) * pvcps(i)
    End If
    Next i
    For i = 0 To UBound(dtps)
    ReDim Preserve convcalc(0 To i)
    If i = 0 Then
        convcalc(i) = ((dtps(i) / 365) ^ 2 + (dtps(i) / 365)) _
        * (pvcps(i) + fv / (1 + dr / 365) ^ dtps(0))
    Else
        convcalc(i) = ((dtps(i) / 365) ^ 2 + (dtps(i) / 365)) * pvcps(i)
    End If
    Next i
    pvfcfs = fv / (1 + dr / 365) ^ dtps(0) + Application.WorksheetFunction.Sum(pvcps)
    durcalcs = Application.WorksheetFunction.Sum(durcalc) / pvfcfs
    convcalcs = Application.WorksheetFunction.Sum(convcalc) / (pvfcfs * (1 + dr) ^ 2)
End If
'Sheet1.Range("$H2").value = pvfcfs
'Sheet1.Range("$I2").value = durcalcs
'Sheet1.Range("$J2").value = convcalcs
ActiveCell.Offset(0, 7).value = pvfcfs
ActiveCell.Offset(0, 8).value = durcalcs
ActiveCell.Offset(0, 9).value = convcalcs
End Sub

'**New function starts here!**

Sub PortfolioMetrics()
Dim row As Long
row = 2
Sheets("bond_portfolio_data").Select
Do Until ActiveSheet.Cells(row, 1) = ""
    ActiveSheet.Cells(row, 1).Select
    Call BondMetrics
    row = row + 1
Loop
ActiveSheet.Cells(row, 8).value = "=SUM(R[-8]C:R[-1]C)"
ActiveSheet.Cells(row, 7).value = "Portfolio Value:"
row = 2
Do Until ActiveSheet.Cells(row, 1) = ""
    ActiveSheet.Cells(row, 1).Select
    ActiveCell.Offset(0, 10).value = "=RC[-3]*RC[-2]"
    row = row + 1
Loop
ActiveSheet.Cells(row, 11).value = "=(SUM(R[-8]C:R[-1]C))/RC[-3]"
ActiveSheet.Cells(row, 10).value = "Portfolio Duration:"
row = 2
Do Until ActiveSheet.Cells(row, 1) = ""
    ActiveSheet.Cells(row, 1).Select
    ActiveCell.Offset(0, 11).value = "=RC[-4]*RC[-2]"
    row = row + 1
Loop
ActiveSheet.Cells(row, 12).value = "=(SUM(R[-8]C:R[-1]C))/RC[-4]"
'ActiveSheet.Cells(row, 11).value = "Portfolio Duration:"
End Sub
