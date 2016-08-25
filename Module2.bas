Option Explicit

Sub ValueBond()

Dim cp As Double, md As Date, ppy As Integer, cpd As Date, cpdMod As Date
Dim cr As Double, br As String, bt As String, wd As Integer, dr As Double
Dim fv As Double, td As Date, cpds() As Variant, i As Integer, integ As Integer
Dim dtp As Long, dtps() As Variant, pvcps() As Variant, pvfcfs As Double
Dim durcalc() As Variant, durcalcs As Double


td = Now
fv = ActiveCell
cr = ActiveCell.Offset(0, 2)
ppy = ActiveCell.Offset(0, 3)
br = ActiveCell.Offset(0, 4)
bt = ActiveCell.Offset(0, 5)
dr = ActiveCell.Offset(0, 6)
If ppy = 0 Then
    cp = 0
Else:
    cp = (fv * cr) / ppy
End If

md = ActiveCell.Offset(0, 1)
If Weekday(md, vbMonday) = 7 Then
    md = md + 1
ElseIf Weekday(md, vbMonday) = 6 Then
    md = md + 2
End If
If ppy = 0 Then
    pvfcfs = fv / (1 + dr / 365) ^ (md - DateValue(CStr(Now())))
    durcalcs = (md - td) / 365
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
    pvfcfs = fv / (1 + dr / 365) ^ dtps(0) + Application.WorksheetFunction.Sum(pvcps)
    durcalcs = Application.WorksheetFunction.Sum(durcalc) / pvfcfs
End If
ActiveCell.Offset(0, 7) = pvfcfs
ActiveCell.Offset(0, 8) = durcalcs
End Sub

