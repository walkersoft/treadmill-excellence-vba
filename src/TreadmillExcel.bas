Attribute VB_Name = "TreadmillExcel"
Option Explicit

Public Const GOAL_DEFINITIONS_TABLE As String = "GoalSetsTable"
Public Const GOAL_SUCCESSES_TABLE As String = "GoalUnlocksTable"
Public Const TREADMILL_LOG_TABLE As String = "MasterDataTable"

Private TreadmillLogData As ListObject
Private GoalData As ListObject
Private GoalSuccesses As ListObject

Public Sub AddTreadmillLogData(activityDate As Date, distance As Single, time As Single, calories As Integer, steps As Integer)
    Application.ScreenUpdating = False
    
    'Check if master data is protected and unprotect it
    Dim isProtected As Boolean
    isProtected = MasterDataSheet.ProtectContents
    If isProtected = True Then MasterDataSheet.Unprotect
    
    'Adds an entry to the master data log from user input
    'Incoming data should be treated as verified and only
    'needing formatting for display purposes as needed
    Dim nextRow As Integer
    Set TreadmillLogData = MasterDataSheet.ListObjects(TREADMILL_LOG_TABLE)
    nextRow = TreadmillLogData.HeaderRowRange.Row + 1
    
    If Not TreadmillLogData.DataBodyRange Is Nothing Then
        nextRow = nextRow + TreadmillLogData.DataBodyRange.Rows.Count
    End If
    
    TreadmillLogData.ListRows.Add
    MasterDataSheet.Range("A" & nextRow).Value = activityDate
    MasterDataSheet.Range("B" & nextRow).Value = Format(distance, "0.00")
    MasterDataSheet.Range("C" & nextRow).Value = Format(time, "0.00")
    MasterDataSheet.Range("D" & nextRow).Value = calories
    MasterDataSheet.Range("E" & nextRow).Value = steps
    
    Call PopulateGoalAchievements
    Call RefreshPivotCache
    
    're-enable master data protection if it was previously set
    If isProtected = True Then MasterDataSheet.Protect
    
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshPivotCache()
    Dim cache As ChartObject
    
    ThisWorkbook.RefreshAll
    'ActiveChart.PivotLayout.PivotTable.PivotFields("Months").Orientation = xlHidden
    'ActiveChart.PivotLayout.PivotTable.PivotFields("Date").AutoGroup
    
    For Each cache In Dashboard.ChartObjects
        With cache.Chart.PivotLayout.PivotTable
            .PivotFields("Months").Orientation = xlHidden
            .PivotFields("Date").AutoGroup
        End With
    Next cache
End Sub

Public Sub PopulateGoalAchievements()
    'Check the treadmill log and goal sets log - exit the sub if either are empty
    Set GoalData = Range(GOAL_DEFINITIONS_TABLE).ListObject
    If GoalData.DataBodyRange Is Nothing Then Exit Sub
    
    Set TreadmillLogData = Range(TREADMILL_LOG_TABLE).ListObject
    If TreadmillLogData.DataBodyRange Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    
    'Clear out the current data in the goal achievements
    Set GoalSuccesses = Range(GOAL_SUCCESSES_TABLE).ListObject
    If Not GoalSuccesses.DataBodyRange Is Nothing Then GoalSuccesses.DataBodyRange.Delete
    
    Dim goalPace As Single
    Dim goalDate As Date
    Dim goalDistance As Single
    Dim startRow As Integer
    Dim goal As Range
    
    'Go through each goal (starting with the newest), and pass the goal data to the
    'goal achievement creation function.
    startRow = TreadmillLogData.DataBodyRange.Rows.Count
    For Each goal In GoalData.DataBodyRange.Rows
        goalDate = goal.Cells(1, 1)
        goalPace = goal.Cells(1, 4)
        goalDistance = goal.Cells(1, 2)
        startRow = CreateGoalEntries(startRow, goalDate, goalPace, goalDistance)
    Next goal
    
    Application.ScreenUpdating = True
End Sub

Private Function CreateGoalEntries(startRow As Integer, goalDate As Date, goalPace As Single, goalDistance As Single) As Long
    'Creates entries in the goal achievment log using goal data passed in.
    'This function receives a row number to start with in the treadmill log
    'and returns the row it stopped on.
    Dim logRow As ListRow
    Dim rowId As Integer
    Dim distance As Single
    Dim time As Single
    Dim pace As Single
    Dim activityDate As Date
    Dim i As Integer
    
    'Calculate where in the achievments table the loop will begin writing
    i = 1
    Set GoalSuccesses = Range(GOAL_SUCCESSES_TABLE).ListObject
    If Not GoalSuccesses.DataBodyRange Is Nothing Then
        i = GoalSuccesses.DataBodyRange.Rows.Count + 1
    End If
    
    Set TreadmillLogData = Range(TREADMILL_LOG_TABLE).ListObject
    For rowId = startRow To TreadmillLogData.HeaderRowRange.Row + 1 Step -1
        'Set the current row as the return value
        CreateGoalEntries = rowId
        
        'Exit the function if the date of the log record is before the date the current goal was set
        activityDate = TreadmillLogData.ListRows(rowId).Range(0, 1).Value
        If activityDate < goalDate Then Exit Function
        
        'Get some data from the log row
        Set logRow = TreadmillLogData.ListRows(rowId)
        distance = logRow.Range(0, 2)
        time = logRow.Range(0, 3)
        pace = time / distance
        
        'Compare goal data to the log data and create an achievment entry if the goal was met
        If (distance >= goalDistance And pace <= goalPace) Then
            GoalSuccesses.ListRows.Add
            GoalSuccesses.ListRows(i).Range(1, 1) = activityDate
            GoalSuccesses.ListRows(i).Range(1, 2) = distance
            GoalSuccesses.ListRows(i).Range(1, 3) = time
            GoalSuccesses.ListRows(i).Range(1, 4).Formula = "=[@Minutes]/[@Miles]"
            i = i + 1
        End If
        
    Next rowId
    
End Function

Public Sub AddGoalSet(activityDate As Date, distance As Double, time As Double)
    Dim GoalTable As ListObject
    Dim Recalculate As VbMsgBoxResult
    
    Set GoalTable = Range(GOAL_DEFINITIONS_TABLE).ListObject
    GoalTable.ListRows.Add 1
    GoalTable.ListRows(1).Range(1, 1) = activityDate
    GoalTable.ListRows(1).Range(1, 2) = distance
    GoalTable.ListRows(1).Range(1, 3) = time
    GoalTable.ListRows(1).Range(1, 4).Formula = "=[@Minutes]/[@Miles]"
    
    Recalculate = MsgBox("Would you like to recalculate goal achievements now?", vbYesNo, "Recalculate?")
    
    If Recalculate = vbYes Then
        Call PopulateGoalAchievements
    End If
End Sub

Public Sub ToggleMasterDataEditing()
    Dim frame As TextFrame
    Dim shp As Shape
    Dim tbl As ListObject
    
    Set tbl = MasterDataSheet.ListObjects("MasterDataTable")
    Set frame = MasterDataSheet.Shapes("ToggleMasterDataEditingButton").TextFrame
    Set shp = MasterDataSheet.Shapes("ToggleMasterDataEditingButton")
    
    If MasterDataSheet.ProtectContents = False Then
        frame.Characters.Text = "Enable Master" + vbNewLine + "Data Editing"
        tbl.TableStyle = "TableStyleMedium4"
        shp.ShapeStyle = msoShapeStylePreset39
        MasterDataSheet.Protect
    Else
        MasterDataSheet.Unprotect
        tbl.TableStyle = "TableStyleMedium2"
        frame.Characters.Text = "Disable Master" + vbNewLine + "Data Editing"
        shp.ShapeStyle = msoShapeStylePreset37
    End If
        
End Sub


Public Sub LoadTreadmillEntryForm()
    DataEntryForm.Show
End Sub

Public Sub LoadGoalEntryForm()
    GoalCreationForm.Show
End Sub
