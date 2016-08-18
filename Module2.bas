Attribute VB_Name = "Module2"
Option Explicit
Sub ValueBond()

Dim cp As Double, md As Date, ppy As Integer, cpd As Date, cpdMod As Date
Dim cr As Double, br As String, bt As String, wd As Integer, dr As Double
Dim fv As Double, td As Date, cpds() As Variant, i As Integer, integ As Integer
Dim dtp As Long, dtps() As Variant, pvcps() As Variant, pvfcfs As Double

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
'Debug.Print md
If ppy = 0 Then
    pvfcfs = fv / (1 + dr / 365) ^ (md - DateValue(CStr(Now())))
Else:
    cpd = DateAdd("M", -12 / ppy, md)
    If Weekday(cpd, vbMonday) = 7 Then
        cpd = cpd + 1
    ElseIf Weekday(cpd, vbMonday) = 6 Then
        cpd = cpd + 2
    End If
    'Debug.Print cpd
    ReDim Preserve cpds(0 To 1)
    cpds(0) = md
    cpds(1) = cpd
    integ = 2
    
    Do
        ReDim Preserve cpds(0 To integ)
        cpd = DateAdd("M", -12 / ppy, cpd)
        'Debug.Print cpd
            If Weekday(cpd, vbMonday) = 7 Then
                cpdMod = cpd + 1
                cpds(integ) = cpdMod
            ElseIf Weekday(cpd, vbMonday) = 6 Then
                cpdMod = cpd + 2
                cpds(integ) = cpdMod
            Else
                cpds(integ) = cpd
            End If
        'Debug.Print cpd
        integ = integ + 1
        
    Loop Until cpd < td
    'Debug.Print UBound(cpds)
    ReDim Preserve cpds(0 To UBound(cpds) - 1)
    'Debug.Print UBound(cpds)
    For i = 0 To UBound(cpds)
        'Debug.Print cpds(i)
        ReDim Preserve dtps(0 To i)
        dtps(i) = cpds(i) - DateValue(CStr(Now()))
        'Debug.Print dtps(i)
    Next i
    
    For i = 0 To UBound(dtps)
        ReDim Preserve pvcps(0 To i)
        pvcps(i) = cp / (1 + dr / 365) ^ dtps(i)
    Next i
    pvfcfs = fv / (1 + dr / 365) ^ dtps(0) + Application.WorksheetFunction.Sum(pvcps)
End If
'Debug.Print (pvfcfs)
ActiveCell.Offset(0, 7) = pvfcfs
End Sub

