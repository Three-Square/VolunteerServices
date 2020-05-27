Attribute VB_Name = "MonthlyRpt2016"
Option Explicit

Dim ReportMonth As String

Sub Master_ConvertToTable()
Attribute Master_ConvertToTable.VB_ProcData.VB_Invoke_Func = " \n14"
    
Dim Master_LastColumn, Master_LastRow As Integer

Sheets("Master").Activate

' Find last column
    Range("A1").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Master_LastColumn = ActiveCell.Column

' Find last row
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Master_LastRow = ActiveCell.Row

    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(Master_LastRow, Master_LastColumn)), , xlYes).Name = _
        "Master"

End Sub

Sub Service_ConvertToTable()
    
Dim Service_LastColumn, Service_LastRow As Integer

Sheets("Service").Activate

' Find last column
    Range("A1").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Service_LastColumn = ActiveCell.Column

' Find last row
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Service_LastRow = ActiveCell.Row

    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(Service_LastRow, Service_LastColumn)), , xlYes).Name = _
        "Service"

End Sub


Sub Service_CalculateDuration()
Attribute Service_CalculateDuration.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Range("Service[[#Headers],[Hours]]").Offset(0, 1).Value = "Duration"
    Range("Service[Duration]").FormulaR1C1 = _
        "=IF(ISERROR(24*([@[To time]]-[@[From time]])),[@[Hours]],24*([@[To time]]-[@[From time]]))"

End Sub

Sub Service_CalculateVisits()

    Range("Service[[#Headers],[Duration]]").Offset(0, 1).Value = "Visits"
    Range("Service[Visits]").FormulaR1C1 = _
        "=IF([@[Duration]]=0,0,[@[Hours]]/[@[Duration]])"

End Sub


Sub Service_JoinKind()

    Range("Service[[#Headers],[Visits]]").Offset(0, 1).Value = "Visit Type"
    Range("Service[Visit Type]").FormulaR1C1 = _
        "=IFERROR(INDEX(Master,MATCH([@Number],Master[Number],0),MATCH(""Kind"",Master[#Headers],0)),"""")"

End Sub


Sub Master_CalculateFirstVisit()

    ReportMonth = InputBox("Enter the number of the month for which you would like to determine first visits.", "Reporting Month")

    Range("Master[[#Headers],[Start date]]").Offset(0, 1).Value = "First Visit"
    Range("Master[First Visit]").FormulaR1C1 = _
        "=IFERROR(IF(MONTH(DATEVALUE([@[Start date]]))=0,""Yes"",""""),"""")"
    Range("Master[First Visit]").Select
    Selection.Replace What:="0", Replacement:=ReportMonth, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub

Sub MonthlyReport()
    
    Sheets.Add
    ActiveSheet.Name = "Monthly Report"
    
    Range("A1").Value = "Volunteer Monthly Report for " & MonthName(ReportMonth)
    Selection.Font.Bold = True
    
    Range("A3").Value = "First Time Volunteers (Individuals):"
    Range("B3").FormulaR1C1 = _
        "=COUNTIFS(Master[First Visit],""=Yes"",Master[Kind],""=Individual"")"
    Range("A4").Value = "First Time Volunteers (Individuals Within Groups):"
    Range("A5").Value = "Total First Time Volunteers (Individuals + Individuals Within Groups):"
    Range("B5").Formula = "=SUM(B3,B4)"
    Range("A6").Value = "First Time Volunteers (Groups):"
    Range("B6").FormulaR1C1 = _
        "=COUNTIFS(Master[First Visit],""=Yes"",Master[Kind],""=Group"")"
    
    Range("A8").Value = "Total Visits (Individuals):"
    Range("B8").FormulaR1C1 = _
        "=SUMIF(Service[Visit Type],""=Individual"",Service[Visits])"
    Range("A9").Value = "Total Visits (Individuals Within Groups):"
    Range("B9").FormulaR1C1 = _
        "=SUMIF(Service[Visit Type],""=Group"",Service[Visits])"
    Range("A10").Value = "Total Visits (Individuals + Individuals Within Groups):"
    Range("B10").Formula = _
        "=SUM(B8,B9)"
    
    Range("A12").Value = "Total Hours of Service (Individuals + Groups)"
    Range("B12").FormulaR1C1 = _
        "=SUM(Service[Hours])"
    
    Range("A3:A12").Select
    With Selection
        .HorizontalAlignment = xlRight
        .Font.Italic = True
        Columns.AutoFit
    End With
    
    Range("B11").Select
    
End Sub

Sub RUN_ME_Volunteer_Monthly_Report()

Application.Run "Master_ConvertToTable"
Application.Run "Master_CalculateFirstVisit"
Application.Run "Service_ConvertToTable"
Application.Run "Service_CalculateDuration"
Application.Run "Service_CalculateVisits"
Application.Run "Service_JoinKind"
Application.Run "MonthlyReport"

End Sub
