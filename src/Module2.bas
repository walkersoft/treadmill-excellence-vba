Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("MasterDataTable[[#Headers],[Date]]").Select
    ActiveSheet.ListObjects("MasterDataTable").TableStyle = "TableStyleMedium4"
    ActiveSheet.ListObjects("MasterDataTable").TableStyle = "TableStyleMedium2"
End Sub
