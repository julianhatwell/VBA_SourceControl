Attribute VB_Name = "TaskListCode"
Const wksName As String = "Tasklist"

Private Sub Refresh()
Attribute Refresh.VB_Description = "Refresh the events"
Attribute Refresh.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveWindow.FreezePanes = False

    With Worksheets(wksName)
        .Range("B3").ListObject.QueryTable.Refresh
    End With
    
End Sub

Private Sub Format_And_Filter()

'frm_WeekNumberInput.Show
TasklistFormat

End Sub

Sub TasklistFormat() '(ByVal week_number As Integer)

    Columns("B:B").ColumnWidth = 10
    Columns("E:E").ColumnWidth = 10
    Columns("F:F").ColumnWidth = 12
    Columns("G:G").ColumnWidth = 12
    Columns("H:H").ColumnWidth = 10
    Columns("J:J").ColumnWidth = 20
    Columns("K:K").ColumnWidth = 20
    Columns("L:L").ColumnWidth = 20
    Columns("M:M").ColumnWidth = 30
    Columns("N:N").ColumnWidth = 10
    
    'ActiveSheet.ListObjects("Table_Kaplan_Scheduler_Tasklist").Range.AutoFilter _
    Field:=4, Criteria1:=week_number
    
    ActiveSheet.ListObjects("Table_Kaplan_Scheduler_Tasklist").Range.AutoFilter _
        Field:=15, Criteria1:="="

    ActiveWorkbook.Worksheets(wksName).ListObjects( _
        "Table_Kaplan_Scheduler_Tasklist").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(wksName).ListObjects( _
        "Table_Kaplan_Scheduler_Tasklist").Sort.SortFields.Add Key:=Range( _
        "Table_Kaplan_Scheduler_Tasklist[event_week]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(wksName).ListObjects( _
        "Table_Kaplan_Scheduler_Tasklist").Sort.SortFields.Add Key:=Range( _
        "Table_Kaplan_Scheduler_Tasklist[event_day]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, CustomOrder:= _
        "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday", DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(wksName).ListObjects( _
        "Table_Kaplan_Scheduler_Tasklist").Sort.SortFields.Add Key:=Range( _
        "Table_Kaplan_Scheduler_Tasklist[event_start_time]"), SortOn:=xlSortOnValues _
        , Order:=xlAscending, DataOption:=xlSortNormal
    'ActiveWorkbook.Worksheets(wksName).ListObjects(
    '    "Table_Kaplan_Scheduler_Tasklist").Sort.SortFields.Add Key:=Range( _
    '    "Table_Kaplan_Scheduler_Tasklist[MID_name]"), SortOn:=xlSortOnValues, Order _
    '    :=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(wksName).ListObjects( _
        "Table_Kaplan_Scheduler_Tasklist").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    Call FormatTime("event_start_time")
    Call FormatTime("event_end_time")
    
    Rows("3:3").Select
    ActiveWindow.FreezePanes = True
    
End Sub

Sub FormatTime(Col As String)

ActiveWorkbook.Worksheets(wksName).ListObjects( _
        "Table_Kaplan_Scheduler_Tasklist").ListColumns( _
        Col).DataBodyRange.NumberFormat = "[$-10409]h:mm:ss AM/PM;@"

End Sub
