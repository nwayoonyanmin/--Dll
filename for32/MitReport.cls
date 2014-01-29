VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public MDB As String
Public RPT As String
Public PrintTo As Byte
Public Formula As String
Public PrinterDirect As Boolean
Function ShowReport()
    Dim A As Integer
    A = DoReport(MDB, RPT, PrintTo, Formula, , PrinterDirect)
End Function


