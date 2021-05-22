Attribute VB_Name = "Common"
Option Explicit

Private Const DATE_ERROR_MESSAGE As String = "Please enter a valid activity date."

Public Function ValidActivityDate(month As String, day As String, year As String) As Boolean
    'Attempts to validate the data in the date inputs is an actual date
    'There is a bit of unexpected behavior as noted below

    If Not IsNumeric(month) Or Not IsNumeric(day) Or Not IsNumeric(year) Then
        Call ShowDateError
        ValidActivityDate = False
        Exit Function
    End If
    
    'IsDate() is "forgiving" and not based on locality; see: https://stackoverflow.com/questions/50108064/why-does-30-9-2013-not-fail-isdate-if-system-date-format-set-for-mm-dd-yyyy
    'something like 30/11/2021 will pass this test when the expectation is that it shouldn't
    If Not IsDate(BuildDateValueFromFormInputs(month, day, year)) Then
        Call ShowDateError
        ValidActivityDate = False
        Exit Function
    End If
    
    ValidActivityDate = True
End Function

Public Function BuildDateValueFromFormInputs(month As String, day As String, year As String) As String
    BuildDateValueFromFormInputs = month & "/" & day & "/" + year
End Function

Public Sub ShowInputError(message As String)
    MsgBox message, vbInformation, "Invalid Data"
End Sub

Public Sub ShowDateError()
    Call ShowInputError(DATE_ERROR_MESSAGE)
End Sub
