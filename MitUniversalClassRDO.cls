VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitUniversalClassRDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim t() As String
Dim N() As Double
Dim D() As Date

Dim TF() As String
Dim NF() As String
Dim DF() As String

Dim pConString As String
Dim DefWhere As String

Public ConString As String
Public BConString As String
Public ActiveConnection As rdo.rdoConnection
Public UseActiveConnection As Boolean

Public ServerName As String
Public DatabaseName As String
Public UserID As String
Public PassWord As String
Public ConnectionType As mitConnTypeCnn

Public DoDebug As Boolean
Public NoOfText As Byte
Public NoOfNumeric As Byte
Public NoOfDate As Byte
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
Function GetText(i As Byte) As String
   GetText = t(i - 1)
End Function
Function GetNumeric(i As Byte) As Double
   GetNumeric = N(i - 1)
End Function
Function GetDate(i As Byte) As Date
   GetDate = D(i - 1)
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
Private Function Writing(rs As rdo.rdoResultset)
    Dim i As Integer
  
    For i = 0 To NoOfText - 1
        rs.rdoColumns(TF(i)) = t(i)
    Next
    For i = 0 To NoOfNumeric - 1
        rs.rdoColumns(NF(i)) = N(i)
    Next
    For i = 0 To NoOfDate - 1
        rs.rdoColumns(DF(i)) = D(i)
    Next
End Function
Private Function Reading(rs As rdo.rdoResultset)
    Dim i As Integer
    
        For i = 0 To NoOfText - 1
            t(i) = rs.rdoColumns(TF(i))
        Next
        For i = 0 To NoOfNumeric - 1
            N(i) = rs.rdoColumns(NF(i))
        Next
        For i = 0 To NoOfDate - 1
            D(i) = IIf(IsNull(rs.rdoColumns(DF(i))), CDate("01/01/1900"), rs.rdoColumns(DF(i)))
        Next
End Function
Private Sub FieldNames()
Dim i As Integer
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
    
    ReDim t(NoOfText)
    ReDim N(NoOfNumeric)
    ReDim D(NoOfDate)
    ReDim TF(NoOfText)
    ReDim NF(NoOfNumeric)
    ReDim DF(NoOfDate)
    If UseActiveConnection Then
    Else
        Select Case ConnectionType
            Case 0
                pConString = ConString
            Case 1 ' Access 97
                pConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                            "DBQ=" & DatabaseName & ";pwd=" & PassWord
            Case 2  'SQL 6.5 Type I
                pConString = "Server=" & ServerName & ";UID=" & UserID _
                      & ";PWD=" & PassWord & ";Database=" & DatabaseName _
                      & ";Driver={SQL Server}"
            Case 3  'SQL 6.5 Type II
                pConString = "Provider=sqloledb;" & _
                    "Data Source=;Initial Catalog=;Server=" & ServerName & ";User Id=" & UserID & _
                    ";Password=" & PassWord & ";Database=" & DatabaseName & ""
            Case 10 ' Oracle
            Case 20 ' DB2
            Case 30 ' Sybase
            Case 40 '
        End Select
    End If
    If DoDebug Then
        MsgBox pConString
    End If

'ServerName
'DatabaseName
'UserID
'PassWord

End Function
Function DefaultFieldNames()
    Dim i As Integer
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
Dim SQL As String
    On Error GoTo er
    
    Dim Cnn As New rdo.rdoConnection
    Dim rdoRs As rdoResultset
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
        Cnn.Connect = pConString
        Cnn.EstablishConnection rdDriverNoPrompt, False
        Cnn.Execute "Set DateFormat dmy"
    End If
    
    SQL = "SELECT max(" + FieldName + ") as M" & _
     " FROM " & TabName & DefaultCondition
       
   Set rdoRs = Cnn.OpenResultset(SQL)
        
        
   If DoDebug Then
        MsgBox SQL
    End If
   
    With rdoRs
        If .EOF Then
            MaxNo = 0
        Else
            MaxNo = !m
        End If
    End With
        
    rdoRs.Close
    Cnn.Close
    
    Exit Function
er:
    MsgBox "MaxNo -" & err.Description
    MaxNo = 0
End Function
Function SaveIt() As Boolean
    Dim SQL As String
    
    On Error GoTo er
    
    Dim Cnn As New rdo.rdoConnection
    Dim rdoRs As rdoResultset
    
    If DoDebug Then
        MsgBox pConString
    End If
    
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
                 
        Cnn.Connect = pConString
        Cnn.EstablishConnection rdDriverNoPrompt, False
        Cnn.Execute "Set DateFormat dmy"
    End If
    
    Set rdoRs = Cnn.OpenResultset(SQL)
   
    SQL = " SELECT * FROM " & TabName & _
                        SearchCondition
    If DoDebug Then
        MsgBox SQL
    End If
    
    
    With rdoRs
    If .EOF Then
        .AddNew
    Else
    End If
        Writing rdoRs
    .Update
    End With
    
    SaveIt = True
    Exit Function
er:
    MsgBox "SaveIt -" & err.Description
    SaveIt = 99
End Function
Function InsertIt() As Byte
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
    MsgBox "InsertIt -" & err.Description
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
    MsgBox "UpdateIt -" & err.Description
    UpdateIt = 99
End Function
Function ReadIt() As Byte
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
        Reading rs
        ReadIt = 255 'True
    Else
        ReadIt = 0 'False
    End If
    rs.Close
    Exit Function
er:
    MsgBox "ReadIt -" & err.Description
    ReadIt = 99 'Error
End Function
Function MoveNext() As Byte
    On Error GoTo er
    Dim Cnn As New rdo.rdoConnection
    Dim rdoRs As rdoResultset
    Dim SQL As String
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
         
        Cnn.Connect = pConString
        Cnn.EstablishConnection rdDriverNoPrompt, False
        Cnn.Execute "Set DateFormat dmy"
    End If
    
    SQL = " SELECT * FROM " & TabName & _
    DefaultCondition & " Order By " & OrderBy
       
   Set rdoRs = Cnn.OpenResultset(SQL)
    
    If DoDebug Then
        MsgBox SQL
    End If
    
    With rdoRs
        If Not .EOF Then
           .MoveNext
           If Not .EOF() Then
              Reading rdoRs
              MoveNext = 255 'True
           Else
              MoveNext = 0 'False
           End If
        Else
            MoveNext = 0 'False
        End If
    End With
    
    rdoRs.Close
    Cnn.Close
    
    Exit Function
er:
    MsgBox "MoveNext -" & err.Description
    MoveNext = 99 'Error
End Function
Function MovePrevious() As Byte
    On Error GoTo er
    
    Dim Cnn As New rdo.rdoConnection
    Dim rdoRs As rdoResultset
    Dim SQL As String
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
         
        Cnn.Connect = pConString
        Cnn.EstablishConnection rdDriverNoPrompt, False
        Cnn.Execute "Set DateFormat dmy"
    End If
    
    SQL = " SELECT * FROM " & TabName & _
    DefaultCondition & " Order By " & OrderBy
       
   Set rdoRs = Cnn.OpenResultset(SQL)
    
    If DoDebug Then
        MsgBox SQL
    End If
    
    With rdoRs
        If Not .EOF Then
           .MovePrevious
           If Not .BOF() Then
              Reading rdoRs
              MovePrevious = 255 'True
           Else
              MovePrevious = 0 'False
           End If
        Else
            MovePrevious = 0 'False
        End If
    End With
    
    rdoRs.Close
    Cnn.Close
    Exit Function
er:
    MsgBox "MoveNext -" & err.Description
    MovePrevious = 99 'Error
End Function
Function MoveLast() As Byte
    On Error GoTo er
    
    Dim Cnn As New rdo.rdoConnection
    Dim rdoRs As rdoResultset
    Dim SQL As String
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
         
        Cnn.Connect = pConString
        Cnn.EstablishConnection rdDriverNoPrompt, False
        Cnn.Execute "Set DateFormat dmy"
    End If
    
    SQL = " SELECT * FROM " & TabName & _
    DefaultCondition & " Order By " & OrderBy
       
   Set rdoRs = Cnn.OpenResultset(SQL)
    
    If DoDebug Then
        MsgBox SQL
    End If
    
    With rdoRs
        If Not .EOF Then
           .MoveLast
            Reading rdoRs
            MoveLast = 255 'True
        Else
            MoveLast = 0 'False
        End If
    End With
    
    rdoRs.Close
    Cnn.Close
    Exit Function
er:
    MsgBox "MoveLast -" & err.Description
    MoveLast = 99 'Error
End Function
Function MoveFirst() As Byte
    On Error GoTo er
    
    Dim Cnn As New rdo.rdoConnection
    Dim rdoRs As rdoResultset
    Dim SQL As String
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
         
        Cnn.Connect = pConString
        Cnn.EstablishConnection rdDriverNoPrompt, False
        Cnn.Execute "Set DateFormat dmy"
    End If
    
    SQL = " SELECT * FROM " & TabName & _
    DefaultCondition & " Order By " & OrderBy
       
   Set rdoRs = Cnn.OpenResultset(SQL)
    
    If DoDebug Then
        MsgBox SQL
    End If
    
    With rdoRs
        If Not .EOF Then
           .MoveFirst
            Reading rdoRs
            MoveFirst = 255 'True
        Else
            MoveFirst = 0 'False
        End If
    End With
    
    rdoRs.Close
    Cnn.Close
    
    Exit Function
er:
    MsgBox "MoveFirst -" & err.Description
    MoveFirst = 99 'Error
End Function
Function FoundIt() As Boolean
    Dim SQL As String
    
    On Error GoTo er
    
    Dim rdoRs As rdoResultset
    Dim Cnn As New rdo.rdoConnection
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
         
        Cnn.Connect = pConString
        Cnn.EstablishConnection rdDriverNoPrompt, False
        Cnn.Execute "Set DateFormat dmy"
    End If
    
    SQL = " SELECT * FROM " & TabName & _
                        SearchCondition
    If DoDebug Then
        MsgBox SQL
    End If
        
        
    With rdoRs
        If Not .EOF Then
            FoundIt = True
        Else
            FoundIt = False
        End If
    End With
    
    rdoRs.Close
    Cnn.Close
    
    Exit Function
er:
    MsgBox "FoundIt -" & err.Description
    FoundIt = False
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
    MsgBox "DeleteIt -" & err.Description
    DeleteIt = 99 'error
End Function
'********************************************************
Function InsertSQL() As Byte
    On Error GoTo er
    Dim Cnn As New ADODB.Connection
    Dim sFields As String
    Dim sValues As String
If DoDebug Then
        MsgBox pConString
End If
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
         
        Cnn.Open pConString
        
    End If
    Cnn.Execute "SET DATEFORMAT DMY"
    '********************
    Dim i As Integer
    
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
        sValues = sValues & IIf(sValues = "", "", ", ") & "'" & D(i - 1) & "'"
    Next
    '*********************
    Dim s As String
    s = "Insert into " & TabName & " (" & sFields & ") " & _
                "Values (" & sValues & ")" & " ; "
    If DoDebug Then
        MsgBox s
    End If
    Cnn.Execute s, i
    InsertSQL = IIf(i = 1, 255, 0)
    Exit Function
er:
    MsgBox "InsertSQL -" & err.Description
    InsertSQL = 99
End Function
Function UpdateSQL() As Byte
    On Error GoTo er
    Dim Cnn As New ADODB.Connection
    Dim sValues As String
    Dim dchr As String
    
    Select Case ConnectionType
           Case 1
                dchr = "#"
           Case Else
                dchr = "'"
    End Select
    
    If UseActiveConnection Then
        Set Cnn = ActiveConnection
    Else
         
        Cnn.Open pConString
    End If
    Cnn.Execute "SET DATEFORMAT DMY"
    '********************
    Dim sFields  As String
    Dim i As Integer
    
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
        "=" & dchr & "" & D(i - 1) & "" & dchr & ""
    Next
    '*********************
    Dim s As String
    
    s = "Update " & TabName & " Set " & sFields & " " & SearchCondition
    If DoDebug Then
        MsgBox s
    End If
    
    Cnn.Execute s, i
    UpdateSQL = IIf(i = 1, 255, 0)
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
    Dim i As Long
    Cnn.Execute s, i
    ExecuteSQL = IIf(i = 1, 255, 0)
    Exit Function
er:
    MsgBox "ExecuteSQL -" & err.Description
    ExecuteSQL = 99
End Function
Function QuerySQL(ss As String) As Variant
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
    cmd.CommandText = ss
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
    MsgBox "QuerySQL -" & err.Description
    QuerySQL = 0
End Function

Public Function Browse(DataType As String, ConnectionOrFile As String, _
    SqlString As String, ReturnKeys As String, _
    Optional FormName, Optional FormWidth, Optional FormHeight, _
    Optional Colheading, Optional ColWidth, Optional ColFormat, _
    Optional ComboList, Optional ComboFields, Optional ComboFieldTypes, _
    Optional DefaultWhere, Optional OrderBy, Optional InitialPosition) As Variant
    If DataType = "A" Then
        If ConnectionOrFile = "" Then
            ConnectionOrFile = DatabaseName
        End If
    Else
        If ConnectionOrFile = "" Then
            ConnectionOrFile = BConString
        End If
    End If
    Browse = DoBrowse(DataType, ConnectionOrFile, _
                SqlString, ReturnKeys, FormName, FormWidth, FormHeight, _
                Colheading, ColWidth, ColFormat, ComboList, ComboFields, _
                ComboFieldTypes, DefaultWhere, OrderBy, InitialPosition, DoDebug)
    
End Function


