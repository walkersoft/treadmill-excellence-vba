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

Sub ClearGoalTables()
    'Deletes all rows in the master data table
    Dim tbl As ListObject
    Dim r As Long
    
    Set tbl = Dashboard.ListObjects(GOAL_SUCCESSES_TABLE)
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    
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

Private Sub gettoggler()
    
    Debug.Print MasterDataSheet.Shapes.Count
    
    For Each s In MasterDataSheet.Shapes
        Debug.Print s.TextFrame.Characters.Text
    Next s
    
End Sub


Sub toggledateedit()
    Dim frame As TextFrame
    Dim shp As Shape
    Dim tbl As ListObject
    
    Set tbl = MasterDataSheet.ListObjects("MasterDataTable")
    Debug.Print "hi"
    Set frame = MasterDataSheet.Shapes("ToggleMasterDataEditingButton").TextFrame
    Set shp = MasterDataSheet.Shapes("ToggleMasterDataEditingButton")
    
    If MasterDataSheet.ProtectContents = False Then
        frame.Characters.Text = "Enable Master" + vbNewLine + "Data Editing"
        tbl.TableStyle = "TableStyleMedium4"
        MasterDataSheet.Protect
    Else
        MasterDataSheet.Unprotect
        tbl.TableStyle = "TableStyleMedium2"
        frame.Characters.Text = "Disable Master" + vbNewLine + "Data Editing"
    End If
        
End Sub
