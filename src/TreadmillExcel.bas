Attribute VB_Name = "TreadmillExcel"
Option Explicit

Public Const GOAL_SETS_DATA_NAME As String = "GoalSetsTable"
Public Const GOAL_UNLOCKS_DATA_NAME As String = "GoalUnlocksTable"
Public Const MASTER_DATA_NAME As String = "MasterDataTable"

Public Sub AddSessionToMasterData(activityDate As Date, distance As Single, time As Single, calories As Integer, steps As Integer)
    'Adds an entry to the master data log from user input
    'Incoming data should be treated as verified and only
    'needing formatting for display purposes as needed
    Dim logData As ListObject
    Dim nextRow As Long
    Set logData = MasterDataSheet.ListObjects(MASTER_DATA_NAME)
    nextRow = logData.HeaderRowRange.Row + 1
    
    If Not logData.DataBodyRange Is Nothing Then
        nextRow = nextRow + logData.DataBodyRange.Rows.Count
    End If
    
    logData.ListRows.Add
    Range("A" & nextRow).Value = activityDate
    Range("B" & nextRow).Value = Format(distance, "0.00")
    Range("C" & nextRow).Value = Format(time, "0.00")
    Range("D" & nextRow).Value = calories
    Range("E" & nextRow).Value = steps
    
End Sub

Private Sub PopulateGoalAchievements()
    
    'Look at the goals/log table for data, no goals/log means no point
    'in continuing, so just exit the sub in that case
    Dim goals As ListObject
    Dim logData As ListObject
    Set goals = Dashboard.ListObjects(GOAL_SETS_DATA_NAME)
    Set logData = MasterDataSheet.ListObjects(MASTER_DATA_NAME)
    If goals.DataBodyRange Is Nothing Then Exit Sub
    If logData.DataBodyRange Is Nothing Then Exit Sub
    
    'Next go through each entry in master data and check for goal
    'achievements. Use the date a goal was set to compare with the
    'dates of entries.
    Dim achievements As ListObject
    Dim goalCount As Integer
    Dim currentGoalPace As Single
    Dim currentGoal As String
    
    Set achievements = Dashboard.ListObjects(GOAL_UNLOCKS_DATA_NAME)
    goalCount = goals.DataBodyRange.Rows.Count
    'currentGoal = goals.DataBodyRange.Cells(1, 1)
    'currentGoalPace = goals.DataBodyRange.Cells(1, 4)
    Debug.Print currentGoal
    
    Dim bottom As Integer
    Dim top As Integer
    Dim nextRow As Integer
    'Debug.Print logData.DataBodyRange.Address
    
    bottom = logData.DataBodyRange.Rows.Count '+ logData.HeaderRowRange.Row
    top = logData.HeaderRowRange.Row + 1
    Dim g As Integer
    Dim currentGoalDist As Single
    
    For g = 1 To goalCount
        currentGoal = goals.DataBodyRange.Cells(g, 1)
        currentGoalPace = goals.DataBodyRange.Cells(g, 4)
        currentGoalDist = goals.DataBodyRange.Cells(g, 2)
        bottom = ProcessLogSegment(bottom, top, currentGoal, currentGoalPace, currentGoalDist)
    Next g
    Debug.Print bottom
End Sub

Private Function ProcessLogSegment(rowStart As Integer, max As Integer, currentGoal As String, currentGoalPace As Single, currentGoalDist As Single) As Long
    'Starting from a given row number, work backwards through the master data log
    'looking for entries that match the given distance/pace amounts. Stop when the log
    'row begin evaluated is no longer on or after the date the given goal was set
    'and return the row number where the processing stopped.
    Dim i As Integer
    Dim logDate As Date
    Dim logData As ListObject
    Dim goalData As ListObject
    Set logData = MasterDataSheet.ListObjects(MASTER_DATA_NAME)
    Set goalData = Dashboard.ListObjects(GOAL_UNLOCKS_DATA_NAME)
    Dim j As Integer
    j = 1
    
    If Not goalData.DataBodyRange Is Nothing Then
        j = goalData.DataBodyRange.Rows.Count + 1
    End If
    
    For i = rowStart To max Step -1
        logDate = logData.ListRows(i).Range(0, 1).Value
        Debug.Print "Log Date: " & logDate & " >= " & currentGoal
        ProcessLogSegment = i
        If logDate < currentGoal Then Exit Function
        
        Dim ldr As ListRow
        Set ldr = logData.ListRows(i)
        Dim ldrPace As Single
        ldrPace = ldr.Range(0, 3) / ldr.Range(0, 2)
        'Debug.Print "yep!"
        If (ldr.Range(0, 2) >= currentGoalDist And ldrPace <= currentGoalPace) Then
            goalData.ListRows.Add
            goalData.ListRows(j).Range(1, 1) = ldr.Range(0, 1)
            goalData.ListRows(j).Range(1, 2) = ldr.Range(0, 2)
            goalData.ListRows(j).Range(1, 3) = ldr.Range(0, 3)
            goalData.ListRows(j).Range(1, 4).Formula = "=[@Minutes]/[@Miles]"
            j = j + 1
            
        End If
        
    Next i
End Function

Public Sub LoadTreadmillEntryForm()
    DataEntryForm.Show
End Sub
