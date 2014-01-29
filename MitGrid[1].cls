VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public DbName As String
Public DataCtrl As Data
Public DBGrd As DBGrid
Public Code As String
Public TabName As String
Public Cols As Integer

'ZMT Added
'**********************************
Public ServerName As String
Public DatabaseName As String
Public ServerTbName As String
Public UserID As String
Public PassWord As String

Public ConnectionType As mitConnTypeCnn

Public SQLSelString As String
'***********************************

Private BeforeEditPosition As Integer
Private pFieldName(100) As Variant
Private pValue(100) As Variant
Private pDataType(100) As Integer
Private pLength(100) As Integer
Private pFormat(100) As Variant
Private pDefault(100) As Variant
Private pWidth(100) As Variant
Private pHeading(100) As Variant
' Class Properties
Public Function SetValue(c As Integer, v As Variant)
    pValue(c) = v
End Function
Public Function SetFieldName(c As Integer, v As Variant)
    pFieldName(c) = v
End Function
Public Function Edit()
    On Error GoTo E
    For i = 0 To Cols
     pValue(i) = IIf(IsNull(DataCtrl.Recordset.Fields(i)), 0, DataCtrl.Recordset.Fields(i))
    Next
    GetPos
E:
End Function
Public Function GetPos()
    On Error Resume Next
    BeforeEditPosition = 0
    BeforeEditPosition = DataCtrl.Recordset.Fields(0)
End Function
Public Function Update()
    On Error Resume Next
    DataCtrl.Recordset.Edit
    For i = 1 To Cols
        DataCtrl.Recordset.Fields(i) = pValue(i)
    Next
    DataCtrl.Recordset.Update
    DataCtrl.Refresh
    GoPos
End Function
Public Function GoPos()
    MoveTo BeforeEditPosition
End Function

Public Property Get Value(c As Integer)
    Value = pValue(c)
End Property
Public Function SetDataType(c As Integer, t As DAO.DataTypeEnum)
    pDataType(c) = t
End Function
Public Function SetDefault(c As Integer, t As Variant)
    pDefault(c) = t
End Function
Public Function SetWidth(c As Integer, t As Integer)
    pWidth(c) = t
End Function
Public Function SetFormat(c As Integer, t As Variant)
    pFormat(c) = t
End Function
Public Function SetHeading(c As Integer, t As Variant)
    pHeading(c) = t
End Function
' Internal Functions
Public Function InitDefaults()
    For i = 1 To 100
        pDefault(i) = ""
    Next
End Function
Public Function InitFieldNames()
    For i = 1 To 100
        pFieldName(i) = "C" + Trim(Str(i))
    Next
End Function
Public Function InitVales()
    For i = 1 To 100
        'pValue(i) = Default(i)
    Next
End Function
' Class Methods/
Public Function AddLine(Optional X As Variant)
    On Error GoTo er
    Dim RecNo As Integer
    If IsMissing(X) Then
        If DataCtrl.Recordset.EOF() And DataCtrl.Recordset.EOF() Then
            RecNo = 1
        Else
            DataCtrl.Recordset.MoveLast
            RecNo = DataCtrl.Recordset!LineNo + 1
        End If
    Else
        RecNo = X
    End If
    DataCtrl.Recordset.AddNew
    For i = 0 To Cols
        Select Case i
            Case 0
                DataCtrl.Recordset!LineNo = RecNo
            Case Else
                DataCtrl.Recordset.Fields(i) = pDefault(i)
        
        End Select
    Next
    DataCtrl.Recordset.Update
    'DataCtrl.Refresh
    FormatGrd
    DataCtrl.Recordset.MoveLast
    DBGrd.Col = 1
    Exit Function
er:
    MsgBox "Addline Error"
End Function
Public Function Bind()
    On Error GoTo er
    DataCtrl.DatabaseName = DbName
    DataCtrl.RecordSource = "select * from [" & TabName & "]"
    DataCtrl.Refresh
    Exit Function
er:
    MsgBox "Bind Error"
End Function
Public Function CreateTemp()
    On Error GoTo er
    Dim db As DAO.Database
    Dim Tb As TableDef
    Dim Fd As DAO.Field
    Dim Ind As Index
    TabName = "T" + Code + Format(Now, "YYYYMMDDHHMMSS")
    Set db = Workspaces(0).OpenDatabase(DbName, True)
    Set Tb = db.CreateTableDef(TabName)
    
    Set Fd = Tb.CreateField("LineNo", dbInteger)
    
    Fd.Required = False
    Tb.Fields.Append Fd
    For i = 1 To Cols
        Select Case pDataType(i)
        Case dbDouble
            Set Fd = Tb.CreateField(pFieldName(i), dbDouble)
            Fd.Required = False
            Tb.Fields.Append Fd
        Case dbInteger
            Set Fd = Tb.CreateField(pFieldName(i), dbInteger)
            Fd.Required = False
            Tb.Fields.Append Fd
        Case dbLong
            Set Fd = Tb.CreateField(pFieldName(i), dbLong)
            Fd.Required = False
            Tb.Fields.Append Fd
        Case dbDate
            Set Fd = Tb.CreateField(pFieldName(i), dbDate)
            Fd.Required = False
            Tb.Fields.Append Fd
        Case dbByte
            Set Fd = Tb.CreateField(pFieldName(i), dbByte)
            Fd.Required = False
            Tb.Fields.Append Fd
        Case dbBoolean
            Set Fd = Tb.CreateField(pFieldName(i), dbBoolean)
            Fd.Required = False
            Tb.Fields.Append Fd
        Case Else   '0, dbText
            Set Fd = Tb.CreateField(pFieldName(i), dbText, 150)
            Fd.Required = False
            Fd.AllowZeroLength = True
            Tb.Fields.Append Fd
        End Select

    Next
    Set Fd = Tb.CreateField("LineNo")
    Set Ind = Tb.CreateIndex("PrimaryKey")
    Ind.Primary = True
    Ind.Unique = True
    Ind.Fields.Append Fd
    Tb.Indexes.Append Ind
    db.TableDefs.Append Tb
    db.Close
    Exit Function
er:
    MsgBox "Create Error"
End Function
Public Function Rows()
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = Workspaces(0).OpenDatabase(DbName)
    Set rs = db.OpenRecordset("Select Count( [LineNo]) as Cnt from [" & TabName + "]")
    rs.MoveLast
    Rows = rs!Cnt
    rs.Close
    db.Close
    Exit Function
er:
    MsgBox "Count Rows Error"
End Function
Public Function SumUp(f As String) As Double
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
    Set rs = db.OpenRecordset("Select sum(" + f + ") as sm from [" & TabName + "]")
    'rs.MoveLast
    SumUp = rs!sm
    rs.Close
    db.Close
    Exit Function
er:
    SumUp = 0
End Function
Public Function TempToRCPT() As Boolean
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
    
    db.Execute "Update [" + TabName + "] Set PayT_ID='01' Where (IsNull(PayT_ID) Or len(trim(PayT_ID))=0) "
    db.Execute "Update [" + TabName + "] Set CurrCode='KS' Where (IsNull(CurrCode) Or len(trim(CurrCode))=0) "
           
     s = "insert into POS_RCP  " & _
     "(Pstation,InvNumber,InvDate,LineNo,PayT_ID,CardNo,Amount,CurrCode,CurrRate,TransType,rStatus) " & _
     " select " & _
     "Pstation,InvNumber,InvDate,LineNo,PayT_ID,CardNo,Amount,CurrCode,CurrRate,TransType,rStatus " & _
     " from " & TabName
    db.Execute "delete * from POS_Rcp"
    db.Execute "INSERT INTO LOG_TMPr SELECT * FROM " & TabName
    db.Execute s
    db.Close
    TempToRCPT = True
    Exit Function
er:
    TempToRCPT = False
End Function
Public Function RCPTtoTemp() As Boolean
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
     s = "insert into " & TabName & _
     " (Pstation,InvNumber,InvDate,LineNo,PayT_ID,CardNo,Amount,CurrCode,CurrRate,TransType,rStatus) " & _
     " select " & _
     "Pstation,InvNumber,InvDate,LineNo,PayT_ID,CardNo,Amount,CurrCode,CurrRate,TransType,rStatus " & _
     " from POS_RCP"
    'db.Execute "delete * from POS_Rcp"
    db.Execute s
    db.Close
    Items2.Refresh
    RCPTtoTemp = True
    
    Exit Function
er:
    RCPTtoTemp = False
End Function
Public Function TempToPOS() As Boolean
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
     s = "insert into POS_DTL  " & _
     "(Pstation,InvNumber,InvDate,LineNo,BarCode,StockID,Description,Measure,Qty,Price,Cost,DiscPercent,Location,DeptNo,StkLevel,Returns,OriginalPrice,UrgentType) " & _
     " select " & _
     "Pstation,InvNumber,InvDate,LineNo,BarCode,StockID,Description,Measure,Qty,Price,Cost,DiscPercent,Location,DeptNo,StkLevel,Returns,OriginalPrice,UrgentType " & _
     " from " & TabName
    db.Execute "Insert into LOG_TMP select * from " & TabName
    db.Execute "delete * from POS_Dtl"
    db.Execute s
    db.Close
    TempToPOS = True
    Exit Function
er:
    TempToPOS = False
End Function
Public Function POSToTemp() As Boolean
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
     s = "insert into " & TabName & "  " & _
     "(Pstation,InvNumber,InvDate,LineNo,BarCode,StockID,Description,Measure,Qty,Price,Cost,DiscPercent,Location,DeptNo,StkLevel,Returns,OriginalPrice,UrgentType) " & _
     " select " & _
     "Pstation,InvNumber,InvDate,LineNo,BarCode,StockID,Description,Measure,Qty,Price,Cost,DiscPercent,Location,DeptNo,StkLevel,Returns,OriginalPrice,UrgentType " & _
     " from POS_DTL"
    db.Execute "Delete * from " & TabName
    db.Execute s
    db.Close
    POSToTemp = True
    Exit Function
er:
    POSToTemp = False
End Function
Public Function UpdateValue(f As String, v As Variant)
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
     s = " Update [" & TabName & "]" & _
                " Set " & f & " = " _
                & v
    db.Execute s
    db.Close
    FormatGrd
    Exit Function
er:
    MsgBox "Update Value Error ! " & err.Description
End Function
Public Function Truncate2(f As String)
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
     s = " Update [" & TabName & "]" & _
                " Set " & f & " = " _
                & "val(format(" & f & ", '0.00'))"
    db.Execute s, adLockOptimistic
    db.Close
    Exit Function
er:
    MsgBox "Truncate2 Up Error"
End Function
Public Function Multiply(f1 As String, f2 As String, f3 As String)
    On Error GoTo er
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = Workspaces(0).OpenDatabase(DbName)
     s = " Update [" & TabName & "]" & _
                " Set " & f1 & " = " _
                & f2 & " * " & f3
    db.Execute s
    'Set rs = db.OpenRecordset("Select sum( [" + f + "]) as sm from [" & TabName + "]")
    'rs.MoveLast
    'SumUp = rs!sm
    'rs.Close
    db.Close
    Exit Function
er:
    MsgBox "Multiply Up Error"
End Function
Public Function DeleteLine()
    On Error GoTo er
    Dim Line As Double
    Line = 0
    If Not DataCtrl.Recordset.EOF() Then
        Line = DataCtrl.Recordset!LineNo
        DataCtrl.Recordset.Delete
        DataCtrl.Refresh
    End If
    If Not DataCtrl.Recordset.EOF() Then
        SortNo
        Cond = "[LineNo]=" & Line - 1 & ""
        If Line - 1 > 0 Then ' Moving back to where it was before
        DataCtrl.Recordset.FindFirst Cond

        End If
        
    End If
    FormatGrd
    Exit Function
er:
    MsgBox "Delete Error"
End Function
Public Function DeleteTable()
    On Error GoTo er
    Dim db As DAO.Database
    DataCtrl.Database.Close
    Set db = Workspaces(0).OpenDatabase(DbName)
    db.Execute "Delete * from [" & TabName & "]"
    db.Close
    Exit Function
er:
    MsgBox "Delete Error"
End Function
Public Function Drop()
    On Error GoTo er
    Dim db As DAO.Database
    DataCtrl.Database.Close
    Set db = Workspaces(0).OpenDatabase(DbName)
    db.Execute "Drop Table  [" & TabName & "]"
    db.Close
    Exit Function
er:
    MsgBox "Drop Error"

End Function
Public Function FormatGrd()
    On Error GoTo er
    For i = 0 To Cols
        DBGrd.Columns(i).Caption = IIf(pHeading(i) = "", DBGrd.Columns(i).Caption, pHeading(i))
        If pDataType(i) = vbDouble Or _
           pDataType(i) = vbInteger Or _
           pDataType(i) = vbDate Then
            DBGrd.Columns(i).NumberFormat = IIf(pFormat(i) = "", "###0", pFormat(i))
            DBGrd.Columns(i).Width = pWidth(i)
        Else
            DBGrd.Columns(i).Width = pWidth(i)

        End If
        If pWidth(i) = 0 Then
            DBGrd.Columns(i).Visible = False
        End If
    Next
    DBGrd.Columns(0).Locked = True
    Exit Function
er:
    MsgBox "Format Grid"
End Function
Public Function MoveTo(v As Integer)
    On Error GoTo er
    If Not DataCtrl.Recordset.EOF Then
        DataCtrl.Recordset.FindFirst "[LineNo]=" & v
    End If
    Exit Function
er:
    MsgBox "Move to Errror"
End Function
Public Function Insert()
    On Error GoTo er
    Dim CurrentLine As Integer
    Dim Reached As Integer
    Reached = 0
    If DataCtrl.Recordset.EOF Then
        AddLine
    Else
        CurrentLine = DataCtrl.Recordset!LineNo
        DataCtrl.Recordset.MoveLast
        Reached = DataCtrl.Recordset!LineNo
        While Reached >= CurrentLine
            DataCtrl.Recordset.Edit
            DataCtrl.Recordset!LineNo = DataCtrl.Recordset!LineNo + 1
            DataCtrl.Recordset.Update
            DataCtrl.Recordset.MovePrevious
            If DataCtrl.Recordset.BOF Then
                Reached = -1
            Else
                Reached = DataCtrl.Recordset!LineNo
            End If

        Wend
        DataCtrl.Refresh
        AddLine CurrentLine
        MoveTo CurrentLine
    End If
    FormatGrd
    Exit Function
er:
    MsgBox "Insert Error"
End Function
Public Function MoveFirst()
    On Error GoTo er
    If Not DataCtrl.Recordset.EOF Then
        DataCtrl.Recordset.MoveFirst
    End If
    Exit Function
er:
    MsgBox "Move Error"
End Function
Public Function MoveNext()
    On Error GoTo er
    DataCtrl.Recordset.MoveNext
    If DataCtrl.Recordset.EOF Then
        DataCtrl.Recordset.MoveLast
    End If
    
    Exit Function
er:
    MsgBox "Move Error"
End Function
Public Function MoveLast()
    On Error GoTo er
    If Not DataCtrl.Recordset.EOF Then
        DataCtrl.Recordset.MoveLast
    End If
    Exit Function
er:
    MsgBox "Move Error"
End Function
Public Function MovePrevious()
    On Error GoTo er
    DataCtrl.Recordset.MovePrevious
    If DataCtrl.Recordset.BOF Then
        DataCtrl.Recordset.MoveFirst
    End If
    Exit Function
er:
    MsgBox "Move Error"
End Function
Public Function Refresh()
    On Error GoTo er
    Dim ColPos As Integer
    ColPos = DBGrd.Col
    GetPos
    DataCtrl.Refresh
    DBGrd.Col = ColPos
    GoPos
    FormatGrd
    Exit Function
er:
    MsgBox "Refresh Error"
End Function
Public Function SortNo()
    If Not DataCtrl.Recordset.EOF() Then
        DataCtrl.Recordset.MoveFirst
        i = 1
        While Not DataCtrl.Recordset.EOF
            DataCtrl.Recordset.Edit
            DataCtrl.Recordset!LineNo = i
            DataCtrl.Recordset.Update
            i = i + 1
            DataCtrl.Recordset.MoveNext
        Wend
    End If
    DataCtrl.Refresh
    FormatGrd
End Function
Public Function InsToLocal(ByVal LConn As ADODB.Connection, ByRef Rws As Long)
    Dim SQL As String
    
    Select Case ConnectionType
        Case 0
            pConString = ConString
        Case 1 ' Access 97
            pConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & DatabaseName & ";pwd=" & PassWord
            'pConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & DatabaseName
        Case 2  'SQL 6.5 Type I ODBC
                pConString = "ODBC;DRIVER={SQL Server};SERVER=" + "ZMTNTS4003" _
                + ";UID=" + UserID + ";PWD=" + PassWord + _
                ";APP=Visual Basics;Database=" + DatabaseName + _
                ";QueryLog_On=Yes;Time=90000;"
        Case 3  'SQL 6.5 Type II ADO
            pConString = "Provider=sqloledb;" & _
                "Data Source=;Initial Catalog=;Server=" & ServerName & ";User Id=" & UserID & _
                ";Password=" & PassWord & ";Database=" & DatabaseName & ""
            
        Case 10 ' Oracle
        
        Case 20 ' DB2
        Case 30 ' Sybase
        Case 40 '
    End Select
    
     
     SQL = "insert into PDetails in '" & DbName & "' " & _
     SQLSelString
  

    
    If Not DoDebug Then
        MsgBox SQL
    End If
    
    LConn.Execute SQL, Rws
        
End Function
Public Function InsToSQLSvr(ByVal LConn As ADODB.Connection, ByRef Rws As Long)
    Dim SQL As String
    Select Case ConnectionType
        Case 0
            pConString = ConString
            
        Case 1 ' Access 97
            pConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & DatabaseName & ";pwd=" & PassWord
            'pConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & DatabaseName
        Case 2  'SQL 6.5 Type I ODBC
                pConString = "ODBC;DRIVER={SQL Server};SERVER=" + "ZMTNTS4003" _
    + ";UID=" + UserID + ";PWD=" + PassWord + _
    ";APP=Visual Basics;Database=" + DatabaseName + _
    ";QueryLog_On=Yes;Time=90000;"
                      
        Case 3  'SQL 6.5 Type II ADO
            pConString = "Provider=sqloledb;" & _
                "Data Source=;Initial Catalog=;Server=" & ServerName & ";User Id=" & UserID & _
                ";Password=" & PassWord & ";Database=" & DatabaseName & ""
            
        Case 10 ' Oracle
        
        Case 20 ' DB2
        Case 30 ' Sybase
        Case 40 '
    End Select
    
    SQL = "insert into [" & pConString & "]." & ServerTbName & _
    " " & SQLSelString
    
    If Not DoDebug Then
        MsgBox SQL
    End If
    
    LConn.Execute SQL, Rws
        
End Function


