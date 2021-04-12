Attribute VB_Name = "TreadmillExcel"
Option Explicit

Public Const MASTER_DATA_NAME As String = "MasterDataTable"
Public Const MASTER_DATA_ROW_OFFSET As Long = 2

Public Sub AddSessionToMasterData(activityDate As Date, distance As Single, time As Single, calories As Integer, steps As Integer)
    'Adds an entry to the master data log from user input
    Dim logData As ListObject
    Dim nextRow As Long
    Set logData = MasterDataSheet.ListObjects(MASTER_DATA_NAME)
    nextRow = MASTER_DATA_ROW_OFFSET
    
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

Public Sub LoadTreadmillEntryForm()
    DataEntryForm.Show
End Sub
