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

Private DATE_ERROR_MESSAGE As String
Private STARTING_DATE As Date

Private Sub UserForm_Initialize()
    STARTING_DATE = Date
    DATE_ERROR_MESSAGE = "Please enter a valid activity date."
    Call SetDateFields
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub SaveAndClose_Click()
    If Not ValidActivityDate Then Exit Sub
    If Not ValidTreadmillData Then Exit Sub
    
    Call LogTreadmillData
    
    Unload Me
End Sub

Private Sub SaveAndNew_Click()
    If Not ValidActivityDate Then Exit Sub
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
    Dim calories As Integer
    Dim steps As Integer

    activityDate = BuildDateValueFromFormInputs
    distance = tbDistance.Value
    time = tbTime.Value
    calories = VBA.CInt(tbCalories.Value)
    steps = VBA.CInt(tbSteps.Value)
    
    Call AddTreadmillLogData(activityDate, distance, time, calories, steps)
End Sub


Private Sub SetDateFields()
    tbMonth.Value = DatePart("m", STARTING_DATE)
    tbDay.Value = DatePart("d", STARTING_DATE)
    tbYear.Value = DatePart("yyyy", STARTING_DATE)
End Sub

Private Function BuildDateValueFromFormInputs() As String
    BuildDateValueFromFormInputs = tbMonth.Value & "/" & tbDay.Value & "/" + tbYear.Value
End Function

Private Function ValidActivityDate() As Boolean
    'Attempts to validate the data in the date inputs is an actual date
    'There is a bit of unexpected behavior as noted below

    If Not IsNumeric(tbMonth.Value) Or Not IsNumeric(tbDay.Value) Or Not IsNumeric(tbYear.Value) Then
        Call ShowDateError
        ValidActivityDate = False
        Exit Function
    End If
    
    'IsDate() is "forgiving" and not based on locality; see: https://stackoverflow.com/questions/50108064/why-does-30-9-2013-not-fail-isdate-if-system-date-format-set-for-mm-dd-yyyy
    'something like 30/11/2021 will pass this test when the expectation is that it shouldn't
    If Not IsDate(BuildDateValueFromFormInputs()) Then
        Call ShowDateError
        ValidActivityDate = False
        Exit Function
    End If
    
    ValidActivityDate = True
End Function

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

Private Sub ShowInputError(message As String)
    MsgBox message, vbInformation, "Invalid Data"
End Sub

Private Sub ShowDateError()
    Call ShowInputError(DATE_ERROR_MESSAGE)
End Sub
