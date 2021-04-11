VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataEntryForm 
   Caption         =   "Enter Session Metrics"
   ClientHeight    =   3690
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

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub SaveAndClose_Click()
    Dim activityDate As Date
    Dim distance As Single
    Dim time As Single
    Dim calories As Integer
    Dim steps As Integer
    
    activityDate = DateValue(VBA.Now)
    distance = 3
    time = 60
    calories = 350
    steps = 5000
    
    Call AddSessionToMasterData(activityDate, distance, time, calories, steps)
    
End Sub
