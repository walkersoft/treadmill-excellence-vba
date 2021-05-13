VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoalCreationForm 
   Caption         =   "Create Goal"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4140
   OleObjectBlob   =   "GoalCreationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoalCreationForm"
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

Private Sub SaveGoal_Click()
    If Not ValidActivityDate Then Exit Sub
    If Not ValidGoalData Then Exit Sub
    
    Call AddGoalSet(BuildDateValueFromFormInputs, tbDistance.Value, tbTime.Value)
    Call Cancel_Click
    
End Sub

Private Function ValidGoalData() As Boolean
    'Validates all of the various data metric fields from the user
    'Verify everything is a number and pop a message box with an error when it isn't
    
    Dim dataChecks(1 To 2, 1 To 2) As String
    dataChecks(1, 1) = tbDistance.Value
    dataChecks(1, 2) = "Enter a valid distance value." & vbNewLine & "(e.g. 2, 3.1, 4.25, etc.)"
    dataChecks(2, 1) = tbTime.Value
    dataChecks(2, 2) = "Enter a valid time in minutes." & vbNewLine & "(e.g. 20, 35.5, 42.25, etc.)"
    'additional future checks need to make sure the array size is adjusted above so the loop runs everything
    
    Dim i As Byte
    For i = 1 To UBound(dataChecks)
        If Not IsNumeric(dataChecks(i, 1)) Then
            Call ShowInputError(dataChecks(i, 2))
            ValidGoalData = False
            Exit Function
        End If
    Next i
    
    ValidGoalData = True
End Function

'Copied from DataEntryForm - should make this common to avoid more duplication if needed elsewhere
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

Private Function BuildDateValueFromFormInputs() As String
    BuildDateValueFromFormInputs = tbMonth.Value & "/" & tbDay.Value & "/" + tbYear.Value
End Function

Private Sub ShowInputError(message As String)
    MsgBox message, vbInformation, "Invalid Data"
End Sub

Private Sub ShowDateError()
    Call ShowInputError(DATE_ERROR_MESSAGE)
End Sub

Private Sub SetDateFields()
    tbMonth.Value = DatePart("m", STARTING_DATE)
    tbDay.Value = DatePart("d", STARTING_DATE)
    tbYear.Value = DatePart("yyyy", STARTING_DATE)
End Sub

