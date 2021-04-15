Attribute VB_Name = "TreadmillExcel"
Option Explicit

Public Const GOAL_DEFINITIONS_TABLE As String = "GoalSetsTable"
Public Const GOAL_SUCCESSES_TABLE As String = "GoalUnlocksTable"
Public Const TREADMILL_LOG_TABLE As String = "MasterDataTable"

Private TreadmillLogData As ListObject
Private GoalData As ListObject
Private GoalSuccesses As ListObject

Public Sub AddSessionToMasterData(activityDate As Date, distance As Single, time As Single, calories As Integer, steps As Integer)
    'Adds an entry to the master data log from user input
    'Incoming data should be treated as verified and only
    'needing formatting for display purposes as needed
    Dim nextRow As Long
    Set TreadmillLogData = MasterDataSheet.ListObjects(TREADMILL_LOG_TABLE)
    nextRow = TreadmillLogData.HeaderRowRange.Row + 1
    
    If Not TreadmillLogData.DataBodyRange Is Nothing Then
        nextRow = nextRow + TreadmillLogData.DataBodyRange.Rows.Count
    End If
    
    TreadmillLogData.ListRows.Add
    Range("A" & nextRow).Value = activityDate
    Range("B" & nextRow).Value = Format(distance, "0.00")
    Range("C" & nextRow).Value = Format(time, "0.00")
    Range("D" & nextRow).Value = calories
    Range("E" & nextRow).Value = steps
    
End Sub

Private Sub PopulateGoalAchievements()
    
    'Look at the GoalData/log table for data, no GoalData/log means no point
    'in continuing, so just exit the sub in that case
    Set GoalData = Dashboard.ListObjects(GOAL_DEFINITIONS_TABLE)
    If GoalData.DataBodyRange Is Nothing Then Exit Sub
    
    Set TreadmillLogData = MasterDataSheet.ListObjects(TREADMILL_LOG_TABLE)
    If TreadmillLogData.DataBodyRange Is Nothing Then Exit Sub
    
    'Next go through each entry in master data and check for goal
    'achievements. Use the date a goal was set to compare with the
    'dates of entries.
    Dim achievements As ListObject
    Dim goalCount As Integer
    Dim currentGoalPace As Single
    Dim currentGoal As String
    
    Set achievements = Dashboard.ListObjects(GOAL_SUCCESSES_TABLE)
    goalCount = GoalData.DataBodyRange.Rows.Count
    'currentGoal = GoalData.DataBodyRange.Cells(1, 1)
    'currentGoalPace = GoalData.DataBodyRange.Cells(1, 4)
    'Debug.Print currentGoal
    
    Dim bottom As Integer
    Dim top As Integer
    Dim nextRow As Integer
    'Debug.Print TreadmillLogData.DataBodyRange.Address
    
    bottom = TreadmillLogData.DataBodyRange.Rows.Count '+ TreadmillLogData.HeaderRowRange.Row
    top = TreadmillLogData.HeaderRowRange.Row + 1
    Dim g As Integer
    Dim currentGoalDist As Single
    
    For g = 1 To goalCount
        currentGoal = GoalData.DataBodyRange.Cells(g, 1)
        currentGoalPace = GoalData.DataBodyRange.Cells(g, 4)
        currentGoalDist = GoalData.DataBodyRange.Cells(g, 2)
        bottom = ProcessLogSegment(bottom, currentGoal, currentGoalPace, currentGoalDist)
    Next g
    Debug.Print bottom
End Sub

Private Function ProcessLogSegment(startRow As Integer, goalDate As String, goalPace As Single, goalDistance As Single) As Long
    'Starting from a given row number, work backwards through the master data log
    'looking for entries that match the given distance/pace amounts. Stop when the log
    'row begin evaluated is no longer on or after the date the given goal was set
    'and return the row number where the processing stopped.
    Dim logRow As ListRow
    Dim rowId As Integer
    Dim distance As Single
    Dim time As Single
    Dim pace As Single
    Dim activityDate As Date
    Dim i As Integer
    
    'calculate where in the achievments table the loop will begin writing
    i = 1
    Set GoalSuccesses = Range(GOAL_SUCCESSES_TABLE).ListObject
    If Not GoalSuccesses.DataBodyRange Is Nothing Then
        i = GoalSuccesses.DataBodyRange.Rows.Count + 1
    End If
    
    Set TreadmillLogData = Range(TREADMILL_LOG_TABLE).ListObject
    For rowId = startRow To TreadmillLogData.HeaderRowRange.Row + 1 Step -1
        ProcessLogSegment = rowId
        activityDate = TreadmillLogData.ListRows(rowId).Range(0, 1).Value
        
        'exit the function if the date of the log record is before the date the current goal was set
        If activityDate < goalDate Then Exit Function
        
        'get some data from the log row
        Set logRow = TreadmillLogData.ListRows(rowId)
        distance = logRow.Range(0, 2)
        time = logRow.Range(0, 3)
        pace = time / distance
        
        'compare goal data to the log data and create an achievment entry if the goal was met
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

Public Sub LoadTreadmillEntryForm()
    DataEntryForm.Show
End Sub
