Attribute VB_Name = "MonthlyRpt2017"
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

Dim ReportMonth As String

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
    Range("A1").Value = "First Time Volunteers (Individuals):"
    Range("B1").FormulaR1C1 = _
        "=COUNTIFS(Master[First Visit],""=Yes"",Master[Kind],""=Individual"")"
    Range("A2").Value = "First Time Volunteers (Individuals Within Groups):"
    Range("A3").Value = "Total First Time Volunteers (Individuals + Individuals Within Groups):"
    Range("B3").Formula = "=SUM(B1,B2)"
    Range("A4").Value = "First Time Volunteers (Groups):"
    Range("B4").FormulaR1C1 = _
        "=COUNTIFS(Master[First Visit],""=Yes"",Master[Kind],""=Group"")"
    
    Range("A6").Value = "Total Visits (Individuals + Individuals Within Groups)"
    Range("B6").FormulaR1C1 = _
        "=SUM(Service[Visits])"
    
    Range("A8").Value = "Total Hours of Service (Individuals + Groups)"
    Range("B8").FormulaR1C1 = _
        "=SUM(Service[Hours])"
    
End Sub

Sub MonthlyReport_AllMacros()

Application.Run "Master_ConvertToTable"
Application.Run "Master_CalculateFirstVisit"
Application.Run "Service_ConvertToTable"
Application.Run "Service_CalculateDuration"
Application.Run "Service_CalculateVisits"
Application.Run "Service_JoinKind"
Application.Run "MonthlyReport"

End Sub
