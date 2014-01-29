VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitUniversalClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim t() As String
Dim N() As Double
Dim D() As Date

Dim TF() As String
Dim NF() As String
Dim DF() As String

Dim PN() As String
Dim PT() As Integer
Dim PV() As Variant

Dim pConString As String
Public ConString As String

Public cmdTimeout As Long
Public cnnTimeout As Long

Public BConString As String
Public ActiveConnection As ADODB.Connection
Public UseActiveConnection As Boolean

Public ServerName As String
Public DatabaseName As String
Public DatabasePath As String
Public UserID As String
Public PassWord As String


Public ConnectionType As mitConnTypeCnn

Public ShowErr As Boolean
Public DoDebug As Boolean
Public NoOfText As Byte
Public NoOfNumeric As Byte
Public NoOfDate As Byte
Public NoOfPara As Byte

Public TabName As String
Public OrderBy As String
Public DefaultWhere As String
Public SearchWhere As String
'*****************************************
Function LetText(i As Byte, v As Variant)
   t(i - 1) = v
End Function
Function LetNumeric(i As Byte, v As Double)
   N(i - 1) = v
End Function
Function LetDate(i As Byte, v As Date)
   D(i - 1) = v
End Function
Function LetParaName(i As Byte, v As Variant)
   PN(i - 1) = v
End Function
Function LetParaType(i As Byte, v As ADODB.ParameterDirectionEnum)
   PT(i - 1) = v
End Function
Function LetParaVal(i As Byte, v As Variant)
   PV(i - 1) = v
End Function
Function GetText(i As Byte) As String
   GetText = t(i - 1)
End Function
Function GetNumeric(i As Byte) As Double
   GetNumeric = N(i - 1)
End Function
Function GetDate(i As Byte) As Date
   GetDate = D(i - 1)
End Function
Function GetParaName(i As Byte) As String
   GetParaName = PN(i - 1)
End Function
Function GetParaType(i As Byte) As Integer
   GetParaType = PT(i - 1)
End Function
Function GetParaVal(i As Byte) As Variant
   GetParaVal = PV(i - 1)
End Function

'********
Function LetTextField(i As Byte, v As Variant)
   TF(i - 1) = v
End Function
Function LetNumericField(i As Byte, v As Variant)
   NF(i - 1) = v
End Function
Function LetDateField(i As Byte, v As Variant)
   DF(i - 1) = v
End Function
'******************************************************
 Private Function DefaultCondition() As String
    If Trim(DefaultWhere) <> "" Then
        DefaultCondition = " where  " + DefaultWhere
    End If
End Function
 Private Function SearchCondition() As String
    If Trim(DefaultWhere) <> "" Then
        SearchCondition = " where  " + DefaultWhere
        If Trim(SearchWhere) <> "" Then
            SearchCondition = SearchCondition + " and " + DefWhere
        End If
    Else
        If Trim(SearchWhere) <> "" Then
            SearchCondition = " where " + SearchWhere
        End If
    End If
End Function
Private Function Writing(rs As ADODB.Recordset)
    For i = 0 To NoOfText - 1
        rs.Fields(TF(i)) = t(i)
    Next
    For i = 0 To NoOfNumeric - 1
        rs.Fields(NF(i)) = N(i)
    Next
    For i = 0 To NoOfDate - 1
        rs.Fields(DF(i)) = D(i)
    Next
End Function
Private Function Reading(rs As ADODB.Recordset)

        For i = 0 To NoOfText - 1
            t(i) = rs.Fields(TF(i)) & ""
        Next
        
        For i = 0 To NoOfNumeric - 1
            N(i) = Val(rs.Fields(NF(i)) & "")
        Next
        For i = 0 To NoOfDate - 1
            D(i) = IIf(IsNull(rs.Fields(DF(i))), CDate("01/01/1900"), rs.Fields(DF(i)))
        Next
End Function
Private Sub FieldNames()
    For i = 1 To NoOfText
        TF(i - 1) = "T" & Trim(Str(i))
    Next
    For i = 1 To NoOfNumeric
        NF(i - 1) = "D" & Trim(Str(i))
    Next
    For i = 1 To NoOfDate
        DF(i - 1) = "N" & Trim(Str(i))
    Next
End Sub
'**************************************************
Function Initialize()
'    If Date > CDate("01/01/2004") Then
'        MsgBox "License Expired"
'        Exit Function
'    End If
'
    ReDim t(NoOfText)
    ReDim N(NoOfNumeric)
    ReDim D(NoOfDate)
    ReDim TF(NoOfText)
    ReDim NF(NoOfNumeric)
    ReDim DF(NoOfDate)
    ReDim PT(NoOfPara)
    ReDim PN(NoOfPara)
    If DatabasePath = "" Then
        DatabasePath = DatabaseName
    End If
    Select Case ConnectionType
        Case 101 To 199
            pConString = ConString
        Case 1 ' Access 97
            pConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & DatabaseName & ";pwd=" & PassWord
            BConString = ";DATABASE=" & DatabaseName & ";PWD=" & PassWord
        Case 2  'SQL 6.5 Type I
                pConString = "Server=" & ServerName & ";UID=" & UserID _
                      & ";PWD=" & PassWord & ";Database=" & DatabaseName _
                      & ";Driver={SQL Server}"
                BConString = "ODBC;Server=" & ServerName & ";UID=sa;PWD=" & PassWord & ";Database=" & DatabaseName & ";Driver={SQL Server}"
        Case 3  'SQL 6.5 Type II ADO
            pConString = "Provider=sqloledb;" & _
                "Data Source=;Initial Catalog=;Server=" & ServerName & ";User Id=" & UserID & _
                ";Password=" & PassWord & ";Database=" & DatabaseName & ""
            BConString = "ODBC;Server=" & ServerName & ";UID=sa;PWD=" & PassWord & ";Database=" & DatabaseName & ";Driver={SQL Server}"
        Case 10 ' Oracle
        Case 20 ' DB2
        Case 30 ' Sybase
        Case 40 '
        Case 90 'MySQL
            pConString = " Driver={mySQL ODBC 3.51 DRIVER}" _
                & ";Pwd=" & PassWord _
                & ";UID=" & UserID _
                & ";Server=" & ServerName _
                & ";Database=" & DatabaseName
    End Select
    
    If DoDebug Then
        MsgBox pConString
    End If

    
End Function
Function InitSPPara()
'    If Date > CDate("01/01/2004") Then
'        MsgBox "License Expired"
'        Exit Function
'    End If
    
    ReDim PT(NoOfPara)
    ReDim PN(NoOfPara)
    ReDim PV(NoOfPara)
    
    For i = 1 To NoOfPara
        PN(i - 1) = ""
    Next
    For i = 1 To NoOfPara
        PT(i - 1) = 0
    Next
    For i = 1 To NoOfPara
        PV(i - 1) = 0
    Next
    

End Function
Function DefaultFieldNames()
    For i = 1 To NoOfText
        TF(i - 1) = "T" & Trim(Str(i))
    Next
    For i = 1 To NoOfNumeric
        NF(i - 1) = "N" & Trim(Str(i))
    Next
    For i = 1 To NoOfDate
        DF(i - 1) = "D" & Trim(Str(i))
    Next
End Function
Function MaxNo(FieldName As String) As Variant
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        ' pConString = ConString
        Cnn.Open pConString
    End If
    
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = "SELECT max(" + FieldName + ") as M" & _
                        " FROM " & TabName & DefaultCondition
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd
    If rs.EOF Then
        MaxNo = 0
    Else
        MaxNo = Val(rs!m & "")
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "MaxNo -" & err.Description
    End If
    MaxNo = 0
End Function
Function SaveIt() As Boolean
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        ' pConString = ConString
        Cnn.Open pConString
    End If
    
    X = ConnectionType Mod 100

    Select Case X
           Case 0 'Connection String Given from interface
           Case 1 'ACCESS MDB
           Case 2 'SQL Type I
                Cnn.Execute "SET DATEFORMAT DMY"
           Case 3 'SQL Type II
                Cnn.Execute "SET DATEFORMAT DMY"
           Case Else
    End Select
    
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
                        SearchCondition
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd

    If rs.EOF Then
        rs.AddNew
    Else
    End If
    Writing rs
    rs.Update
    rs.Close
    SaveIt = True
    Exit Function
er:
    If ShowErr Then
        MsgBox "SaveIt -" & err.Description
    End If
    SaveIt = 99
End Function
Function InsertIt() As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    If DoDebug Then
        MsgBox pConString
    End If
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
                        SearchCondition
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd
    
    If rs.EOF Then
        rs.AddNew
        Writing rs
        rs.Update
        InsertIt = True
    Else
        InsertIt = False
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "InsertIt - " & err.Description
    End If
    InsertIt = 99
End Function
Function UpdateIt() As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
                        SearchCondition
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    
    rs.Open cmd
    If rs.EOF Then
        UpdateIt = 0 'False
    Else
        Writing rs
        rs.Update
        UpdateIt = 255 'True
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "UpdateIt -" & err.Description
    End If
    UpdateIt = 99
End Function
Function ReadIt(Optional ErMsg) As Byte
    If IsMissing(ErMsg) Then
        ErMsg = True
    End If
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        ' pConString = ConString
        Cnn.Open pConString
    End If
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
    SearchCondition
    
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd

    If Not rs.EOF Then
        Reading rs
        ReadIt = 255 'True
    Else
        ReadIt = 0 'False
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "ReadIt -" & err.Description
    End If
    ReadIt = 99 'Error
End Function
Function MoveNext() As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenDynamic

    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    
'    cmd.CommandText = "SELECT * FROM  " & TabName & _
    " " & SearchCondition & " Order By " & OrderBy
    
    cmd.CommandText = "SELECT * FROM  " & TabName & _
    SearchCondition & " Order By " & OrderBy
    
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
       
    rs.Open cmd

    If Not rs.EOF Then
        rs.MoveNext
        
        If Not rs.EOF Then
            Reading rs
            MoveNext = 255 'True
        Else
            MoveNext = 0 'False
        End If
        Else
            MoveNext = 0 'False
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "MoveNext -" & err.Description
    End If
    MoveNext = 99 'Error
End Function
Function MovePrevious() As Byte
    On Error GoTo er
    
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenDynamic
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
    SearchCondition & " Order By " & OrderBy
    
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MovePrevious
        If Not rs.BOF Then
            Reading rs
            MovePrevious = 255 'True
        Else
            MovePrevious = 0 'False
        End If
        Else
            MovePrevious = 0 'False
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "MoveNext -" & err.Description
    End If
    MovePrevious = 99 'Error
End Function
Function MoveLast() As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenDynamic
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
    DefaultCondition & " Order By " & OrderBy
    
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd
    If Not rs.EOF Then
        rs.MoveLast
        Reading rs
        MoveLast = 255 'True
    Else
        MoveLast = 0 'False
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "MoveLast -" & err.Description
    End If
    MoveLast = 99 'Error
End Function
Function MoveFirst() As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenDynamic
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
    DefaultCondition & " Order By " & OrderBy
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd
    If Not rs.EOF Then
        rs.MoveFirst
        Reading rs
        MoveFirst = 255 'True
    Else
        MoveFirst = 0 'False
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "MoveFirst -" & err.Description
    End If
    MoveFirst = 99 'Error
End Function
Function FoundIt() As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
                        SearchCondition
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    
    rs.Open cmd
    If Not rs.EOF Then
        FoundIt = 255
    Else
        FoundIt = 0
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "FoundIt -" & err.Description
    End If
    FoundIt = 99
End Function
Function DeleteIt() As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = " SELECT * FROM " & TabName & _
                        SearchCondition
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd

    If Not rs.EOF Then
        rs.Delete
        DeleteIt = 255 'true
    Else
        DeleteIt = 0 'false
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "DeleteIt -" & err.Description
    End If
    DeleteIt = 99 'error
End Function
'********************************************************
Function InsertSQL() As Long
    On Error GoTo er
    Dim Cnn As New ADODB.Connection
    Dim errLoop As ADODB.Error
    Dim sFields As String
    Dim sValues As String

    
    If DoDebug Then
            MsgBox pConString
            
            'MsgBox "State " & ActiveConnection.State
    End If
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    X = ConnectionType Mod 100
    If DoDebug Then
        MsgBox "x " & X
        MsgBox "x " & Cnn
        MsgBox "x " & pConString
    End If
    Select Case X
           Case 0 'Connection String Given from interface
           Case 1 'ACCESS MDB
                dchr = "DateValue('"
                dchrEnd = "')"
           Case 2 'SQL Type I
                dchr = "'"
                dchrEnd = "'"
                Cnn.Execute "SET DATEFORMAT DMY"
           Case 3 'SQL Type II
                dchr = "'"
                dchrEnd = "'"
                Cnn.Execute "SET DATEFORMAT DMY"
           Case Else
                dchr = "'"
                dchrEnd = "'"
    End Select

    '********************
    Cnn.CommandTimeout = 30
    
    For i = 1 To NoOfText
        sFields = sFields & IIf(sFields = "", "", ", ") & TF(i - 1)
    Next
    For i = 1 To NoOfNumeric
        sFields = sFields & IIf(sFields = "", "", ", ") & NF(i - 1)
    Next
    For i = 1 To NoOfDate
        sFields = sFields & IIf(sFields = "", "", ", ") & DF(i - 1)
    Next
    '*********************
    For i = 1 To NoOfText
        sValues = sValues & IIf(sValues = "", "", ", ") & "'" & t(i - 1) & "'"
    Next
    For i = 1 To NoOfNumeric
        sValues = sValues & IIf(sValues = "", "", ", ") & "" & N(i - 1) & ""
    Next
    For i = 1 To NoOfDate
        sValues = sValues & IIf(sValues = "", "", ", ") & dchr & D(i - 1) & dchrEnd
    Next
    '*********************
    s = "Insert into " & TabName & " (" & sFields & ") " & _
                "Values (" & sValues & ")" & " ; "
    If DoDebug Then
        MsgBox s
    End If
    Cnn.Execute s, i
    InsertSQL = IIf(i > 0, 255, 0)
    Exit Function
er:
    If ShowErr Then
        MsgBox "InsertSQL - " & err.Description
    End If
    For Each errLoop In Cnn.Errors
         InsertSQL = errLoop.Number
         Exit For
    Next
    
    
'    InsertSQL = 99
End Function
Function UpdateSQL() As Byte
    On Error GoTo er
    Dim Cnn As New ADODB.Connection
    Dim sValues As String
    Dim dchr As String
    Dim dchrEnd As String
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    
    Dim X As Byte
    
    X = ConnectionType Mod 100
    If DoDebug Then
        MsgBox "x " & X
        MsgBox "Cnn " & Cnn
        MsgBox "pCon " & pConString
        MsgBox UseActiveConnection
    End If
    Select Case X
           Case 0 'Connection String Given from interface
           Case 1 'ACCESS MDB
                dchr = "DateValue('"
                dchrEnd = "')"
           Case 2 'SQL Type I
                dchr = "'"
                dchrEnd = "'"
                Cnn.Execute "SET DATEFORMAT DMY"
           Case 3 'SQL Type II
                dchr = "'"
                dchrEnd = "'"
                Cnn.Execute "SET DATEFORMAT DMY"
           Case Else
                dchr = "'"
                dchrEnd = "'"
    End Select
    '********************
    sFields = ""
    For i = 1 To NoOfText
        sFields = sFields & IIf(sFields = "", "", ", ") & TF(i - 1) & _
             "='" & t(i - 1) & "'"
    Next
    For i = 1 To NoOfNumeric
        sFields = sFields & IIf(sFields = "", "", ", ") & NF(i - 1) & _
        "=" & N(i - 1) & ""
    Next
    
    
    For i = 1 To NoOfDate
        sFields = sFields & IIf(sFields = "", "", ", ") & DF(i - 1) & _
        "=" & dchr & "" & D(i - 1) & "" & dchrEnd & ""
    Next
    '*********************
    s = "Update " & TabName & " Set " & sFields & " " & SearchCondition
    If DoDebug Then
        MsgBox s
        'MsgBox "State " & ActiveConnection.State
    End If
    
    Cnn.Execute s, i
    UpdateSQL = IIf(i > 0, 255, 0)
    Exit Function
er:
    MsgBox "UpdateSQL -" & err.Description
    UpdateSQL = 99
End Function
Function ExecuteSQL(s As String) As Byte
    On Error GoTo er
    Dim Cnn As New ADODB.Connection
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    If DoDebug Then
        MsgBox s
    End If
    
    
    Select Case X
           Case 0 'Connection String Given from interface
           Case 1 'ACCESS MDB
           Case 2 'SQL Type I
                Cnn.Execute "SET DATEFORMAT DMY"
           Case 3 'SQL Type II
                Cnn.Execute "SET DATEFORMAT DMY"
           Case Else
    End Select
    
    
    
    Cnn.Execute s, i
    ExecuteSQL = IIf(i > 0, 255, 0)
    Exit Function
er:
    If ShowErr Then
        MsgBox "ExecuteSQL -" & err.Description
    End If
    ExecuteSQL = 99
End Function
Function ExecuteSP(Sp As String) As Byte
    On Error GoTo er
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
        
    If cnnTimeout <= 0 Then
       Cnn.CommandTimeout = 5000
    End If
    
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
            
    If cnnTimeout <= 0 Then
       Cnn.CommandTimeout = 50000
    End If
    
    cmd.CommandText = Sp
    cmd.Name = "SpName"
    cmd.CommandType = adCmdStoredProc
    
    Set cmd.ActiveConnection = Cnn
    
    cmd.Parameters.Refresh
    
    If cmdTimeout <= 0 Then
        cmdTimeout = 50000
    End If
    
    cmd.CommandTimeout = cmdTimeout
    
    For i = 0 To NoOfPara - 1
        cmd.Parameters("@" & PN(i)).Value = PV(i)
        cmd.Parameters("@" & PN(i)).Direction = PT(i)
    Next i
       
    Select Case X
           Case 0 'Connection String Given from interface
           Case 1 'ACCESS MDB
           Case 2 'SQL Type I
                cmd.Execute "SET DATEFORMAT DMY"
           Case 3 'SQL Type II
                cmd.Execute "SET DATEFORMAT DMY"
           Case Else
    End Select
       
       
    cmd.Execute r
    
    For i = 0 To NoOfPara - 1
        PV(i) = cmd.Parameters("@" & PN(i)).Value
        If DoDebug Then
            MsgBox PV(i)
        End If
    Next i
    
    ExecuteSP = IIf(r > 0, 255, 0)
    
    Exit Function
er:
    If ShowErr Then
        MsgBox "ExecuteSQL -" & err.Description
    End If
    ExecuteSP = 99
End Function
Function QuerySQL(ss As String) As Variant
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        ' pConString = ConString
        Cnn.Open pConString
    End If
    cmd.Name = "yourCommandName"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = ss
    
    X = ConnectionType Mod 100

    Select Case X
           Case 0 'Connection String Given from interface
           Case 1 'ACCESS MDB
           Case 2 'SQL Type I
                Cnn.Execute "SET DATEFORMAT DMY"
           Case 3 'SQL Type II
                Cnn.Execute "SET DATEFORMAT DMY"
           Case Else
    End Select
    
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd
    If rs.EOF Then
        QuerySQL = 0
    Else
        QuerySQL = rs!Ret
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "QuerySQL -" & err.Description
    End If
    QuerySQL = 0
End Function
Function SqlQuery(Stmt As String, Ret As Variant) As Byte
    On Error GoTo er
    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    
    X = ConnectionType Mod 100

    Select Case X
           Case 0 'Connection String Given from interface
           Case 1 'ACCESS MDB
           Case 2 'SQL Type I
                Cnn.Execute "SET DATEFORMAT DMY"
           Case 3 'SQL Type II
                Cnn.Execute "SET DATEFORMAT DMY"
           Case Else
    End Select
    
    
    cmd.Name = "SqlQuery"
    cmd.ActiveConnection = Cnn
    cmd.CommandText = Stmt
    
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    rs.Open cmd
    If rs.EOF Then
        SqlQuery = 0
    Else
        SqlQuery = 255
        If IsNull(rs!Ret) Then
            SqlQuery = 0
        Else
            Ret = rs!Ret
        End If
    End If
    rs.Close
    Exit Function
er:
    If ShowErr Then
        MsgBox "SqlQuery -" & err.Description
    End If
    SqlQuery = 99
End Function
Public Function Browse(DataType As String, ConnectionString As String, _
    SqlString As String, ReturnKeys As String, _
    Optional FormName, Optional FormWidth, Optional FormHeight, _
    Optional Colheading, Optional ColWidth, Optional ColFormat, _
    Optional ComboList, Optional ComboFields, Optional ComboFieldTypes, _
    Optional DefaultWhere, Optional OrderBy, Optional InitialPosition) As Variant
        If ConnectionString = "" Then
            If BConString <> "" Then
                ConnectionString = BConString 'DatabaseName
            Else
                ConnectionString = ConString
            End If
        End If
    If DoDebug Then
        MsgBox "Connection String to Server ->" & ConnectionString
    End If
    'If DatabasePath = "" And DataType = "A" Then
    '    DatabasePath = DatabaseName
    'End If
    Browse = DoBrowse(DataType, ConnectionString, _
                SqlString, ReturnKeys, FormName, FormWidth, FormHeight, _
                Colheading, ColWidth, ColFormat, ComboList, ComboFields, _
                ComboFieldTypes, DefaultWhere, OrderBy, InitialPosition, DoDebug, DatabasePath)
    
End Function
Public Function LoadToCombo(c As Variant, TbName As String, _
fldIdx As String, fldDesc As String, Optional FltCond As Variant)

    Dim rs As New ADODB.Recordset
    Dim Cnn As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim Cnt As Byte
    Dim i As Byte
    On Error GoTo er
    
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Open pConString
    End If
    
    cmd.Name = "yourCommandName"
    
    cmd.ActiveConnection = Cnn
    
    If IsMissing(FltCond) Then
       FltCond = " WHERE 1=1"
    Else
       FltCond = " WHERE " & FltCond
    End If
    
    cmd.CommandText = "SELECT * FROM " & TbName & _
    FltCond
    
    If DoDebug Then
        MsgBox cmd.CommandText
    End If
    
    rs.Open cmd
    
    c.Clear
    
    If rs.EOF() Then Exit Function
    
    While Not rs.EOF
    
        c.AddItem rs.Fields(fldDesc)
        On Error Resume Next
        c.ItemData(c.NewIndex) = rs.Fields(fldIdx)
        On Error GoTo er
        rs.MoveNext
        i = i + 1
    Wend
    
    rs.Close
    Exit Function
er:
    aa = 99 'Error
End Function

