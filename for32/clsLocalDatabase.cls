VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mitLocalDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public EncryptType As Byte

Public Description As String
Public DatabasePath As String

Public ServerName As String
Public DatabaseName As String
Public UserID As String
Public UserPassword As String
Public DSNname  As String

Public BranchCode As String
Public BranchName As String

Public AppPath As String
Public DefaultPath As String
Public DefaultPassword As String
Public Function ReadDefault(Optional DefaultType) As Boolean
    On Error GoTo er
    ReadDefault = False
    AppPath = IIf(AppPath = "", App.Path, AppPath)
    DefaultPath = AppPath + "\Default.mdb"
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim Cmd1  As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    dConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DefaultPath & ";pwd=" & DefaultPassword
    Cnn.Open dConString
    Cmd1.Name = "ReadProfile"
    Cmd1.ActiveConnection = Cnn
    Cmd1.CommandText = " SELECT * FROM Profile"
    rs.Open Cmd1
    If Not rs.EOF Then
        DatabasePath = rs!DbPath
        Description = rs!CoName
       
        ReadDefault = True
    Else
        ReadDefault = False
    End If
    
    
    rs.Close
    
    Cnn.Close
    Exit Function
er:
    MsgBox err.Description
    ReadDefault = False
End Function
Public Function ReadDefaultNew(Optional DefaultType) As Boolean
    On Error GoTo er
    ReadDefaultNew = False
    AppPath = IIf(AppPath = "", App.Path, AppPath)
    DefaultPath = AppPath + "\Default.mdb"
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim Cmd1  As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    dConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DefaultPath & ";pwd=" & DefaultPassword
    Cnn.Open dConString
    Cmd1.Name = "ReadProfile"
    Cmd1.ActiveConnection = Cnn
    Cmd1.CommandText = " SELECT * FROM Profile"
    rs.Open Cmd1
    If Not rs.EOF Then
        DatabasePath = rs!DbPath
        Description = rs!CoName
       
        ServerName = rs!ServerName & ""
        DatabaseName = rs!DatabaseName & ""
        UserID = rs!UserID
        UserPassword = rs!Pwd
        ReadDefaultNew = True
    Else
        ReadDefaultNew = False
    End If
    
    
    rs.Close
    
    Cnn.Close
    Exit Function
er:
    MsgBox err.Description
    ReadDefaultNew = False
End Function

Public Function ReadDefaultCM(Optional DefaultType) As Boolean
    On Error GoTo er
    ReadDefaultCM = False
    AppPath = IIf(AppPath = "", App.Path, AppPath)
    DefaultPath = AppPath + "\Default.mdb"
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim Cmd1  As New ADODB.Command
    Dim Cmd2 As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    dConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DefaultPath & ";pwd=" & DefaultPassword
    Cnn.Open dConString
    Cmd1.Name = "ReadProfile"
    Cmd1.ActiveConnection = Cnn
    Cmd1.CommandText = " SELECT * FROM Profile"
    rs.Open Cmd1
    If Not rs.EOF Then
        DatabasePath = rs!DbPath
        Description = rs!CoName
        ReadDefaultCM = True
    Else
        ReadDefaultCM = False
    End If
    rs.Close
    
    Cmd2.Name = "ReadSetServer"
    Cmd2.ActiveConnection = Cnn
    Cmd2.CommandText = " Select * from SetServer "
    rs.Open Cmd2
    If Not rs.EOF Then
        ServerName = rs!LSN & ""
        DatabaseName = rs!LDN & ""
        UserID = rs!LID & ""
        UserPassword = rs!LPW & ""
        BranchCode = rs!BranchNo
        ReadDefaultCM = True
    Else
        ReadDefaultCM = False
    End If
    rs.Close
    
    Cnn.Close
    Exit Function
er:
    MsgBox err.Description
    ReadDefaultCM = False
End Function
Public Function ReadFullyOnline(Optional DefaultType) As Boolean
    On Error GoTo er
    ReadFullyOnline = False
    AppPath = IIf(AppPath = "", App.Path, AppPath)
    DefaultPath = AppPath + "\Setup.mdb"
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim Cmd1  As New ADODB.Command
    Dim Cmd2 As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    dConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DefaultPath & ";pwd=" & DefaultPassword
    Cnn.Open dConString
    Cmd1.Name = "ReadSetup"
    Cmd1.ActiveConnection = Cnn
    Cmd1.CommandText = " SELECT * FROM Setup"
    rs.Open Cmd1
    If Not rs.EOF Then
        BranchCode = rs!BranchCode & ""
        BranchName = rs!BranchName & ""
    
        ServerName = rs!ServerName & ""
        DatabaseName = rs!DbName & ""
        UserID = rs!UserID
        UserPassword = rs!Pwd
        
        ReadFullyOnline = True
    Else
        ReadFullyOnline = False
    End If
    rs.Close
    Cnn.Close
    Exit Function
er:
    MsgBox err.Description
    ReadFullyOnline = False
End Function
Public Function ReadIRBS(Optional DefaultType) As Boolean
    On Error GoTo er
    ReadIRBS = False
    AppPath = IIf(AppPath = "", App.Path, AppPath)
    DefaultPath = AppPath + "\LocalDatabase.mdb"
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim Cmd1  As New ADODB.Command
    Dim Cmd2 As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    dConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DefaultPath & ";pwd=" & DefaultPassword
    Cnn.Open dConString
    Cmd1.Name = "ReadBankDatabases"
    Cmd1.ActiveConnection = Cnn
    Cmd1.CommandText = " SELECT * FROM BankDatabases where isLocal=-1"
    rs.Open Cmd1
    If Not rs.EOF Then
        BranchCode = rs!BranchCode & ""
        BranchName = rs!BranchName & ""
    
        ServerName = rs!ServerName & ""
        DatabaseName = rs!DbName & ""
        UserID = rs!UserID
        UserPassword = rs!Pwd
        
        ReadIRBS = True
    Else
        ReadIRBS = False
    End If
    rs.Close
    Cnn.Close
    Exit Function
er:
    MsgBox err.Description
    ReadIRBS = False
End Function
Public Function ReadIrbsRemote(sLocalConnString As String, BrCode As String) As Boolean
    On Error GoTo er
    ReadIrbsRemote = False
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim Cmd1  As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    Cnn.Open sLocalConnString
    Cmd1.Name = "ReadBankDatabases"
    Cmd1.ActiveConnection = Cnn
    Cmd1.CommandText = " SELECT * FROM BankDatabases where BranchCode='" _
                        & BrCode & "'"
    rs.Open Cmd1
    If Not rs.EOF Then
        BranchCode = rs!BranchCode & ""
        BranchName = rs!BranchName & ""
        ServerName = rs!ServerName & ""
        DatabaseName = rs!DbName & ""
        UserID = rs!UserID
        UserPassword = rs!Pwd
        ReadIrbsRemote = True
    Else
        ReadIrbsRemote = False
    End If
    rs.Close
    Cnn.Close
    Exit Function
er:
    MsgBox err.Description
    ReadIrbsRemote = False
End Function

Public Function ReadGeneral(Optional DbID) As Boolean
    On Error GoTo er
    ReadGeneral = False
    AppPath = IIf(AppPath = "", App.Path, AppPath)
    DefaultPath = AppPath + "\LocalGeneral.mdb"
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim Cmd1  As New ADODB.Command
    Dim Cmd2 As New ADODB.Command
    Dim DatabaseID As String
    
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    dConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DefaultPath & ";pwd=" & DefaultPassword
    Cnn.Open dConString
    Cmd1.Name = "ReadBankDatabases"
    Cmd1.ActiveConnection = Cnn
    If IsMissing(DbID) Then
        DatabaseID = "'0'"
    Else
        DatabaseID = "'" & DbID & "'"
    End If
    Cmd1.CommandText = " SELECT * FROM ServerDatabases where DatabaseID=" & DatabaseID
    rs.Open Cmd1
    If Not rs.EOF Then
        BranchCode = rs!BranchCode & ""
        BranchName = rs!BranchName & ""
    
        ServerName = rs!ServerName & ""
        DatabaseName = rs!DbName & ""
        UserID = rs!UserID
        UserPassword = rs!Pwd
        
        DatabasePath = rs!DbPath & ""
        
        ReadGeneral = True
    Else
        ReadGeneral = False
    End If
    rs.Close
    Cnn.Close
    Exit Function
er:
    MsgBox err.Description
    ReadGeneral = False
End Function
Function PrepareConString(ConnectionType As Byte, Optional BwrConString)
    Select Case ConnectionType Mod 100
        Case 1 ' Access 97
            If DatabasePath = "" Then
                DatabasePath = DatabaseName
            End If
                PrepareConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DatabasePath & ";pwd=" & UserPassword
           ' PrepareConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" _
            & DatabasePath & ";," & UserID & "," & UserPassword

        Case 71 ' Access 97
                PrepareConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & DatabasePath & ";pwd=" & UserPassword
            
            'PrepareConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" _
            & DatabaseName & ";," & UserID & "," & UserPassword
        Case 2  'SQL 6.5 Type I
                PrepareConString = "Server=" & ServerName & ";UID=" & UserID _
                      & ";PWD=" & UserPassword & ";Database=" & DatabaseName _
                      & ";Driver={SQL Server}"
                BConString = "ODBC;Server=" & ServerName & ";UID=" & UserID & ";PWD=" & UserPassword & ";Database=" & DatabaseName & ";Driver={SQL Server}"
        Case 3  'SQL 6.5 Type II ADO
            PrepareConString = "Provider=sqloledb;" & _
                "Data Source=;Initial Catalog=;Server=" & ServerName & ";User Id=" & UserID & _
                ";Password=" & UserPassword & ";Database=" & DatabaseName & ""
            BwrConString = "ODBC;Server=" & ServerName & ";UID=" & UserID & ";PWD=" & UserPassword & ";Database=" & DatabaseName & ";Driver={SQL Server}"
        Case 10 ' Oracle
        
        Case 20 ' DB2
        Case 30 ' Sybase
        Case 40 '
        Case 90 'MySQL
            PrepareConString = " Driver={mySQL ODBC 3.51 DRIVER}" _
                    & ";Pwd=" & UserPassword _
                    & ";UID=" & UserID _
                    & ";Server=" & ServerName _
                    & ";Database=" & DatabaseName
    End Select
End Function



