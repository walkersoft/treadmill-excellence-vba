Attribute VB_Name = "Demo"
Option Explicit

Sub ClearTable()
    'Deletes all rows in the master data table
    Dim tbl As ListObject
    Dim r As Long
    
    Set tbl = MasterDataSheet.ListObjects(MASTER_DATA_NAME)
    
    For r = tbl.DataBodyRange.Rows.Count To 1 Step -1
        tbl.ListRows(r).Delete
    Next r
    
End Sub

Sub AddData()
    Dim activityDate As Date
    Dim distance As Single
    Dim time As Single
    Dim calories As Integer
    Dim steps As Integer
    
    activityDate = dateValue(VBA.Now)
    distance = 3
    time = 60
    calories = 350
    steps = 5000
    
    Call AddSessionToMasterData(activityDate, distance, time, calories, steps)
End Sub
