VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public ShowErr As Boolean
Public ServerName As String
Public DatabaseName As String
Public UserID As String
Public UserPassword As String
Public DSNname  As String
Public DatabasePath As String
Public Function GetConnected(Cnn As ADODB.Connection, ConnString As String, CnnTime As Long, CmmTime As Long) As Boolean
    On Error GoTo err
    Cnn.ConnectionString = ConnString
    Cnn.CursorLocation = adUseClient
    Cnn.ConnectionTimeout = CnnTime
    Cnn.CommandTimeout = CmmTime
    Cnn.Open
    GetConnected = True
    Exit Function
err:
    GetConnected = False
    If ShowErr Then
        MsgBox err.Description
    End If
End Function
Function GetConString(ConnectionType As Byte, Optional BwrConString)
    Select Case ConnectionType
           Case 1 ' Access 97
                GetConString = DatabasePath & ";PWD=" & UserPassword
                BwrConString = DatabasePath & ";PWD=" & UserPassword
           Case 2  'SQL 6.5 Type I
                GetConString = "ODBC;DRIVER={SQL Server};SERVER=" + ServerName _
                + ";UID=" + UserID + ";PWD=" + UserPassword + _
                ";APP=Visual Basics;Database=" + DatabaseName + _
                ";QueryLog_On=Yes;Time=90000"
                BwrConString = "ODBC;Server=" & ServerName & ";UID=" & UserID & ";PWD=" & UserPassword & ";Database=" & DatabaseName & ";Driver={SQL Server}"
        Case 3  'SQL 6.5 Type II ADO
               
                GetConString = "Provider=sqloledb;" & _
                "Data Source=;Initial Catalog=;Server=" & ServerName & ";User Id=" & UserID & _
                ";Password=" & UserPassword & ";Database=" & DatabaseName & ""
                BwrConString = "ODBC;Server=" & ServerName & ";UID=" & UserID & ";PWD=" & UserPassword & ";Database=" & DatabaseName & ";Driver={SQL Server}"
        Case 10 ' Oracle
        
        Case 20 ' DB2
        Case 30 ' Sybase
        Case 40 '
        Case 90 'MySQL
                GetConString = "ODBC;Driver={mySQL ODBC 3.51 DRIVER}" _
                & ";Pwd=" & UserPassword _
                & ";UID=" & UserID _
                & ";Option=131072;" _
                & "Stmt=" _
                & ";Server=" & ServerName _
                & ";Database=" & DatabaseName
    End Select
End Function
Function RegionalDateOK() As Boolean
    If Format(Date, "dd/mm/yyyy") = Str(Date) Then
        RegionalDateOK = True
    Else
        RegionalDateOK = False
    End If
End Function
Function Encrypt(Pwd As String, EncryptType As Byte) As String
    Select Case EncryptType
        Case 0
            Encrypt = Pwd
        Case 1
            Encrypt = EncryptType1(Pwd)
        Case Else
            Encrypt = Pwd
        
    End Select
End Function
Function Decrypt(Pwd As String, EncryptType As Byte) As String
    Select Case EncryptType
        Case 0
            Decrypt = Pwd
        Case 1
            Decrypt = DecryptType1(Pwd)
        Case Else
            Decrypt = Pwd
    End Select
End Function
Function BreakUp(s As Variant, a() As Variant, Optional Sep) As Integer
    If IsMissing(Sep) Then
        Sep = ","
    End If
    Dim StrLen As Integer
    Dim Temp As Variant
    StrLen = Len(s)
    For i = 1 To StrLen
        
        If Mid(s, i, 1) = Sep Or i = StrLen Then
            BreakUp = BreakUp + 1
            ReDim Preserve a(BreakUp)
            a(BreakUp - 1) = Temp
            If i = StrLen Then
                 a(BreakUp - 1) = a(BreakUp - 1) + Mid(s, i, 1)
            End If
            Temp = ""
        Else
            Temp = Temp + Mid(s, i, 1)
        End If
    Next
    
End Function
'****************************************************
Private Function EncryptType1(s As String) As String
    Dim Temp As String
    For i = 1 To Len(s)
        Temp = Temp + EncryptChar1(Mid(s, i, 1))
    Next
    EncryptType1 = Temp
End Function
Private Function DecryptType1(s As String) As String
    Dim Temp As String
    For i = 1 To Len(s)
        Temp = Temp + DecryptChar1(Mid(s, i, 1))
    Next
    DecryptType1 = Temp
End Function
Private Function EncryptChar1(c As String) As String
    X = Asc(c)
    If X Mod 2 <> 0 Then
        EncryptChar1 = Chr(X + 14)
    Else
        EncryptChar1 = Chr(X - 6) 'can div by 2
    End If
End Function
Private Function DecryptChar1(c As String) As String
    X = Asc(c)
    If X Mod 2 <> 0 Then
        DecryptChar1 = Chr(X - 14)
    Else
        DecryptChar1 = Chr(X + 6)
    End If
End Function
'********************************************************

