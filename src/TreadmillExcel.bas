Attribute VB_Name = "TreadmillExcel"
Option Explicit

Public Const MASTER_DATA_TBL As String = "MasterDataTable"
Public Const MASTER_DATA_ROW_OFFSET As Long = 2


Public Sub AddSessionToMasterData(activityDate As Date, distance As Single, time As Single, calories As Integer, steps As Integer)
    'Adds an entry to the master data log from user input
    Dim logData As ListObject
    Dim nextRow As Long
    Set logData = MasterDataSheet.ListObjects(MASTER_DATA_TBL)
    nextRow = MASTER_DATA_ROW_OFFSET
    
    If Not logData.DataBodyRange Is Nothing Then
        nextRow = nextRow + logData.DataBodyRange.Rows.Count
    End If
    
    Debug.Print nextRow
    logData.ListRows.Add
    Range("A" & nextRow).Value = activityDate
    Range("B" & nextRow).Value = distance
    Range("C" & nextRow).Value = time
    Range("D" & nextRow).Value = calories
    Range("E" & nextRow).Value = steps
'    Debug.Print activityDate
'    Debug.Print distance
'    Debug.Print time
'    Debug.Print calories
'    Debug.Print steps
    
End Sub
