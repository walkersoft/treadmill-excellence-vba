VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataEntryForm 
   Caption         =   "Enter Session Metrics"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5250
   OleObjectBlob   =   "DataEntryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private STARTING_DATE As Date

Private Sub UserForm_Activate()
    'This will help position the form better in the application
    'in the event the user has multiple monitors
    Me.top = (Application.Height / 2) - (Me.Width / 2)
    Me.left = (Application.Width / 2) - (Me.Height / 2)
End Sub

Private Sub UserForm_Initialize()
    STARTING_DATE = Date
    Call SetDateFields
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub SaveAndClose_Click()
    If Not ValidActivityDate(tbMonth.Value, tbDay.Value, tbYear.Value) Then Exit Sub
    If Not ValidTreadmillData Then Exit Sub
    
    Call LogTreadmillData
    
    Unload Me
End Sub

Private Sub SaveAndNew_Click()
    If Not ValidActivityDate(tbMonth.Value, tbDay.Value, tbYear.Value) Then Exit Sub
    If Not ValidTreadmillData Then Exit Sub
    
    Call LogTreadmillData
    
    'UX related; reset form fields and setup for next entry
    tbDistance.Value = ""
    tbTime.Value = ""
    tbCalories.Value = ""
    tbSteps.Value = ""
    tbMonth.SetFocus
End Sub

Private Sub LogTreadmillData()
    Dim activityDate As Date
    Dim distance As Single
    Dim time As Single
    Dim calories As Long
    Dim steps As Long

    activityDate = BuildDateValue(tbMonth.Value, tbDay.Value, tbYear.Value)
    distance = tbDistance.Value
    time = tbTime.Value
    calories = VBA.CLng(tbCalories.Value)
    steps = VBA.CLng(tbSteps.Value)
    
    Call AddTreadmillLogData(activityDate, distance, time, calories, steps)
End Sub

Private Sub SetDateFields()
    tbMonth.Value = DatePart("m", STARTING_DATE)
    tbDay.Value = DatePart("d", STARTING_DATE)
    tbYear.Value = DatePart("yyyy", STARTING_DATE)
End Sub

Private Function ValidTreadmillData() As Boolean
    'Validates all of the various data metric fields from the user
    'Verify everything is a number and pop a message box with an error when it isn't
    
    Dim dataChecks(1 To 4, 1 To 2) As String
    dataChecks(1, 1) = tbDistance.Value
    dataChecks(1, 2) = "Enter a valid distance value." & vbNewLine & "(e.g. 2, 3.1, 4.25, etc.)"
    dataChecks(2, 1) = tbTime.Value
    dataChecks(2, 2) = "Enter a valid time in minutes." & vbNewLine & "(e.g. 20, 35.5, 42.25, etc.)"
    dataChecks(3, 1) = tbCalories.Value
    dataChecks(3, 2) = "Enter a valid calorie amount (whole numbers only)." & vbNewLine & "(e.g. 200, 350, etc.)"
    dataChecks(4, 1) = tbSteps.Value
    dataChecks(4, 2) = "Enter a valid step count (whole numbers only)." & vbNewLine & "(e.g. 2500, 3480, etc.)"
    'additional future checks need to make sure the array size is adjusted above so the loop runs everything
    
    Dim i As Byte
    For i = 1 To UBound(dataChecks)
        If Not IsNumeric(dataChecks(i, 1)) Then
            Call ShowInputError(dataChecks(i, 2))
            ValidTreadmillData = False
            Exit Function
        End If
    Next i
    
    ValidTreadmillData = True
End Function
