Attribute VB_Name = "Module1"
Option Explicit

Sub LockTable()
Attribute LockTable.VB_ProcData.VB_Invoke_Func = " \n14"
'
' LockTable Macro
'

'
    Range("MasterDataTable[#All]").Select
    Selection.Locked = True
    Selection.FormulaHidden = False
End Sub
