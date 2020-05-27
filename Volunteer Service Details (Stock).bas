Attribute VB_Name = "VolunteerSvcDetails"
Sub Erase_Zeroes_And_Blanks()
Attribute Erase_Zeroes_And_Blanks.VB_ProcData.VB_Invoke_Func = " \n14"
'
' With_Total_Hours Macro
'

'
    Sheets("Worksheet 1").Select
    Sheets("Worksheet 1").Copy After:=Sheets(1)
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name _
        = "With_Total_Hours"
    Sheets("Worksheet 1 (2)").Select
    Sheets("Worksheet 1 (2)").Name = "With Total Hours"
    ActiveSheet.ListObjects("With_Total_Hours").Range.AutoFilter Field:=5, _
        Criteria1:="=0", Operator:=xlOr, Criteria2:="="
    Range("With_Total_Hours[Hours]").Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("With_Total_Hours").Range.AutoFilter Field:=5
    Range("With_Total_Hours[[#Headers],[Volunteer]]").Select
    
    Sheets("Worksheet 1").Select
    Sheets("Worksheet 1").Copy After:=Sheets(2)
    Sheets("Worksheet 1 (2)").Select
    Sheets("Worksheet 1 (2)").Name = "Without Total Hours"
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name _
        = "Without_Total_Hours"
    ActiveSheet.ListObjects("Without_Total_Hours").Range.AutoFilter Field:=5, Criteria1:= _
        "=0", Operator:=xlOr, Criteria2:="="
    Range("Without_Total_Hours[Hours]").Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Without_Total_Hours").Range.AutoFilter Field:=5
    ActiveSheet.ListObjects("Without_Total_Hours").Range.AutoFilter Field:=2, Criteria1:= _
        "<>"
    ActiveSheet.ListObjects("Without_Total_Hours").Range.AutoFilter Field:=2, Criteria1:="="
    Range("Without_Total_Hours[Service From Date]").Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("Without_Total_Hours").Range.AutoFilter Field:=2
    Range("Without_Total_Hours[[#Headers],[Volunteer]]").Select
End Sub
