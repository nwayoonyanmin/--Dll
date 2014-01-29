VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MITConnection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ServerName As String
Public DatabaseName As String
Public UserID As String
Public PassWord As String

Public Enum mitConnTypeCnn
       mitConnMsAccessExt = 101
       mitConnODBCExt = 102
       mitConnADOExt = 103
       mitConnOracleExt = 110
       mitConnDB2Ext = 120
       mitConnSysbaseExt = 130
       
       mitConnMsAccess = 1
       mitConnODBC = 2
       mitConnADO = 3
       mitConnOracle = 10
       mitConnDB2 = 20
       mitConnSysbase = 30
End Enum

Public ConnectionType  As mitConnTypeCnn
Public cmdTimeout As Long
Public ConString As String
Public DoDebug As Boolean
Public CursorLoc As ADODB.CursorLocationEnum
Dim pConString As String
Dim BConString As String
Function SetConnection(LConn As ADODB.Connection)
    
'    If Date > CDate("01/01/2004") Then
'        MsgBox "License Expired", vbInformation
'        Exit Function
'    End If
'
    Select Case ConnectionType
           Case 0
                pConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DatabaseName & ";pwd=" & PassWord
           Case 1 ' Access 97
              '  pConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DatabaseName & ";pwd=" & PassWord
                pConString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" _
                & DatabaseName & ";Jet OLEDB:Database Password=" & PassWord
'                pConString = ";" & DatabaseName & ";PWD=" & PassWord
                
           Case 2  'SQL 6.5 Type I
                 pConString = "ODBC;DRIVER={SQL Server};SERVER=" + ServerName _
                 + ";UID=" + UserID + ";PWD=" + PassWord + _
                 ";APP=Visual Basics;Database=" + DatabaseName + _
                 ";QueryLog_On=Yes"
                    BConString = "ODBC;Server=" & ServerName & ";UID=" & UserID & ";PWD=" & PassWord & ";Database=" & DatabaseName & ";Driver={SQL Server}"
           Case 3  'SQL 6.5 Type II ADO
                pConString = "Provider=sqloledb;" & _
                    "Data Source=;Initial Catalog=;Server=" & ServerName & ";User Id=" & UserID & _
                    ";Password=" & PassWord & ";Database=" & DatabaseName & ""
                BConString = "ODBC;Server=" & ServerName & ";UID=" & UserID & ";PWD=" & PassWord & ";Database=" & DatabaseName & ";Driver={SQL Server}"
           Case 10 ' Oracle
            
           Case 20 ' DB2
           Case 30 ' Sybase
           Case 40 '
           Case 90
                    pConString = "Driver={mySQL ODBC 3.51 DRIVER}" _
                    & ";Pwd=" & PassWord _
                    & ";UID=" & UserID _
                    & ";Option=131072;" _
                    & "Stmt=" _
                    & ";Server=" & ServerName _
                    & ";Database=" & DatabaseName
'
           '         pConString = "Provider=MySQLProv;" & _
                    "Data Source=" & DatabaseName & ";" & _
                     ";Option=131072;" _
                    & "Stmt=" & _
                    ";Location=" & ServerName & ";User Id=" & UserID & ";" & _
                    "Password=" & PassWord & ""
    End Select
    
    If cmdTimeout = 0 Then
       cmdTimeout = 100
    End If
    
    ConString = pConString
    LConn.CommandTimeout = cmdTimeout
    LConn.Open pConString
'MsgBox LConn
    If DoDebug Then
       MsgBox "State " & LConn.State
       MsgBox pConString
    End If
    
End Function
