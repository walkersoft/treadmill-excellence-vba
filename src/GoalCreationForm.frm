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

Private STARTING_DATE As Date

Private Sub UserForm_Initialize()
    STARTING_DATE = Date
    
    Call SetDateFields
End Sub

Private Sub UserForm_Activate()
    'this will help position the form better in the application
    'in the event the user has multiple monitors
    Me.top = (Application.Height / 2) - (Me.Width / 2)
    Me.left = (Application.Width / 2) - (Me.Height / 2)
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub SaveGoal_Click()
    If Not ValidActivityDate(tbMonth.Value, tbDay.Value, tbYear.Value) Then Exit Sub
    If Not ValidGoalData Then Exit Sub
    
    Call AddGoalSet(BuildDateValueFromFormInputs(tbMonth.Value, tbDay.Value, tbYear.Value), tbDistance.Value, tbTime.Value)
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


Private Sub SetDateFields()
    tbMonth.Value = DatePart("m", STARTING_DATE)
    tbDay.Value = DatePart("d", STARTING_DATE)
    tbYear.Value = DatePart("yyyy", STARTING_DATE)
End Sub

