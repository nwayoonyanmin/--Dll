VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitAdoGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public DbName As String
Public DataCtrl As Adodc
Public DBGrd As DataGrid
Public Code As String
Public TabName As String
Public Cols As Integer

Public TempDbPath  As String
Public TempDbUID As String
Public TempDbPWD As String
Public dFlds As String
Public sFlds As String
Public fltStr As String
'ZMT Added
'**********************************
Public ServerName As String
Public DatabaseName As String
Public ServerTbName As String
Public UserID As String
Public PassWord As String

Public ConnectionType As mitConnTypeCnn

Public SQLSelString As String
Public iSWithIndex As Boolean
Public sIndexfld As String


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
    Dim i As Integer
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
    If Not DataCtrl.Recordset.EOF() Then
       BeforeEditPosition = DataCtrl.Recordset.Fields(0)
    End If
End Function
Public Function Update()
    On Error Resume Next
    Dim i  As Integer
    'DataCtrl.Recordset.Edit
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
Public Function SetDataType(c As Integer, t As ADODB.DataTypeEnum)
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
    Dim i As Integer
    For i = 1 To 100
        pDefault(i) = ""
    Next
End Function
Public Function InitFieldNames()
    Dim i As Integer
    For i = 1 To 100
        pFieldName(i) = "C" + Trim(Str(i))
    Next
End Function
Public Function InitVales()
    Dim i As Integer
    For i = 1 To 100
        'pValue(i) = Default(i)
    Next
End Function
' Class Methods/
Public Function AddLine(Optional X As Variant)
    On Error GoTo er
    Dim RecNo As Integer
    Dim i As Integer
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
    DataCtrl.CursorType = adOpenDynamic
    DataCtrl.CursorLocation = adUseClient
    DataCtrl.LockType = adLockOptimistic
    DataCtrl.Mode = adModeReadWrite
'    MsgBox DataCtrl.ConnectionString
   DataCtrl.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
                        
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
    Dim i As Integer
    
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
Public Function CreateTempAdo()
    On Error GoTo er
    Dim iConn As New ADODB.Connection
    Dim i As Integer
    Dim iConString As String
    Dim sCreateStr As String
    Dim sIndxStr As String
    
  '  iConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" _
     & TempDbPath & ";" & TempDbUID & ";" & TempDbPWD
       iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
          
     
     iConn.Open iConString
     
    TabName = "T" + Code + Format(Now, "YYYYMMDDHHMMSS")


    sCreateStr = "CREATE TABLE " & TabName & "("

    For i = 0 To Cols
        Select Case pDataType(i)
               Case adDouble
                    sCreateStr = sCreateStr & pFieldName(i) & " DOUBLE "
               Case adSmallInt
                    sCreateStr = sCreateStr & pFieldName(i) & " SHORT "
               Case adInteger
                    sCreateStr = sCreateStr & pFieldName(i) & " LONG "
               
               Case adDBDate, adDBTime, adDBTimeStamp
                    sCreateStr = sCreateStr & pFieldName(i) & " DATETIME "
               Case adTinyInt
                    sCreateStr = sCreateStr & pFieldName(i) & " BYTE "
               Case adBoolean
                    sCreateStr = sCreateStr & pFieldName(i) & " BIT "
               Case Else   '0, dbText
                    sCreateStr = sCreateStr & pFieldName(i) & " TEXT "
    
               
        End Select
        
           If i < Cols Then
                   sCreateStr = sCreateStr & ","
               Else
                  sCreateStr = sCreateStr & ")"
          End If
    Next
    
    iConn.BeginTrans
    iConn.Execute sCreateStr
    
       If iSWithIndex Then
          sIndxStr = "CREATE UNIQUE INDEX PrimaryKey " _
          & "ON " & TabName & " (" & sIndexfld & " ASC) " _
          & "WITH PRIMARY;"
              iConn.Execute sIndxStr
       End If

       iConn.CommitTrans



    
    
    iConn.Close
    Set iConn = Nothing
    Exit Function
er:
    iConn.Close
    MsgBox "Create Error", vbInformation
End Function


Public Function Rows()
    On Error GoTo er
    Dim iRs As New ADODB.Recordset
    Dim iConn As New ADODB.Connection
    Dim iConString  As String
    Dim Rws As Long
    Dim selSql As String
    
    'iConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" _
     & TempDbPath & ";," & TempDbUID & "," & TempDbPWD
            
     iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
    iConn.Open iConString
    
    'selSql = "Select Count( [LineNo]) as Cnt from [" & TabName + "]"
    'iRs.Open selSql, iConn, adOpenKeyset, adLockOptimistic
    
    Set iRs = iConn.Execute("Select Count( [LineNo]) as Cnt from [" & TabName + "]", Rws)
    
    'MsgBox iRs!Cnt
    'iRs.MoveLast
    'Rows = Rws ' rs!Cnt
    Rows = iRs!Cnt
    iRs.Close
    iConn.Close
    
    Exit Function
er:
    MsgBox "Count Rows Error"
End Function
Public Function SumUp(fldName As String) As Double
    On Error GoTo er
    
    Dim iRs As New ADODB.Recordset
    Dim iConn As New ADODB.Connection
    Dim iConString As String
    'iConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" _
     & TempDbPath & ";," & TempDbUID & "," & TempDbPWD
            
     iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
    iConn.Open iConString
    
    
    Set iRs = iConn.Execute("Select sum(" + fldName + ") as sm from [" & TabName + "]")
        
  
    If Not iRs.EOF() Then
       SumUp = IIf(IsNull(iRs!sm), 0, iRs!sm)
    Else
       SumUp = 0
    End If
    
    iRs.Close
    iConn.Close
    Exit Function
er:
    SumUp = 0
End Function

Public Function UpdateValue(f As String, v As Variant)
    On Error GoTo er
    On Error GoTo er
    
    Dim iConn As New ADODB.Connection
    Dim iConString As String
    Dim s As String
    
    iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
    iConn.Open iConString
    
    s = " Update [" & TabName & "]" & _
                " Set " & f & " = " _
                & v
    iConn.Execute s
    
    iConn.Close
    
    FormatGrd
    Exit Function
er:
    MsgBox "Update Value Error ! " & err.Description
End Function
Public Function Truncate2(f As String)
    On Error GoTo er
    Dim iConn As New ADODB.Connection
    Dim s As String
    Dim iConString As String
    iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
    iConn.Open iConString

     s = " Update [" & TabName & "]" & _
                " Set " & f & " = " _
                & "val(format(" & f & ", '0.00'))"
    iConn.Execute s, adLockOptimistic
    iConn.Close
    Exit Function
er:
    MsgBox err.Description & "Truncate2 Up Error"
End Function
Public Function Multiply(f1 As String, f2 As String, f3 As String)
    On Error GoTo er
    
    Dim iRs As New ADODB.Recordset
    Dim iConn As New ADODB.Connection
    Dim s As String
    Dim iConString As String
    'iConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" _
     & TempDbPath & ";," & TempDbUID & "," & TempDbPWD
            
    iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
    
    iConn.Open iConString
    s = " Update [" & TabName & "]" & _
                " Set " & f1 & " = " _
                & f2 & " * " & f3
    
    iConn.Execute s
    iConn.Close
    Exit Function
er:
    MsgBox "Multiply Up Error"
End Function
Public Function DeleteLine()
    On Error GoTo er
    Dim Line As Double
    Dim Cond As String
    
    Line = 0
  
    If Not DataCtrl.Recordset.EOF() Then
        Line = DataCtrl.Recordset!LineNo
        DataCtrl.Recordset.Delete adAffectCurrent
    End If
    
    Cond = "LineNo=" & Line - 1
    
    If Not DataCtrl.Recordset.EOF() Then
       FillLineNo
       DataCtrl.Refresh
       If Line - 1 > 0 Then ' Moving back to where it was before
          DataCtrl.Recordset.Find Cond
       Else
       End If
    End If
    Exit Function
er:
    MsgBox "Delete Error." & err.Description, vbInformation
End Function
Public Function DeleteTable()
    On Error GoTo er
    Dim iConn As New ADODB.Connection
    Dim s As String
    Dim iConString As String
    
    iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
    
    iConn.Open iConString

    s = "Delete * from [" & TabName & "]"
    iConn.Execute s
    iConn.Close
    Exit Function
er:
    MsgBox "Delete Error"
End Function
Public Function Drop()
    On Error GoTo er
    Dim iConn As New ADODB.Connection
    Dim s As String
    Dim iConString As String
                
    iConString = "Driver={Microsoft Access Driver (*.mdb)}; " & _
                        "DBQ=" & TempDbPath & ";pwd=" & TempDbPWD
    
    iConn.Open iConString

    s = "Drop Table  [" & TabName & "]"
    iConn.Execute s
    iConn.Close
    Exit Function
er:
    MsgBox "Drop Error"

End Function
Public Function FormatGrd()
    On Error GoTo er
    Dim i As Integer
    For i = 0 To Cols
    
    
        DBGrd.Columns(i).Caption = IIf(pHeading(i) = "", DBGrd.Columns(i).Caption, pHeading(i))
      '  DBGrd.Columns (i)
        Select Case pDataType(i)
               Case adDouble, adCurrency, adSingle
                    DBGrd.Columns(i).NumberFormat = IIf(pFormat(i) = "", "Standard", pFormat(i))
                    DBGrd.Columns(i).Alignment = dbgRight
               Case adInteger, adSmallInt, adNumeric, adTinyInt
                    DBGrd.Columns(i).NumberFormat = IIf(pFormat(i) = "", "#,##0", pFormat(i))
                    DBGrd.Columns(i).Alignment = dbgRight
               Case adDBDate
                    DBGrd.Columns(i).NumberFormat = IIf(pFormat(i) = "", "dd/mm/yyyy", pFormat(i))
                    DBGrd.Columns(i).Alignment = dbgRight
               Case adVarChar
                    DBGrd.Columns(i).NumberFormat = IIf(pFormat(i) = "", "", pFormat(i))
                    DBGrd.Columns(i).Alignment = dbgLeft
               Case Else
                    DBGrd.Columns(i).Alignment = dbgLeft
        End Select
        
        DBGrd.Columns(i).Width = pWidth(i)
        
        

        If pWidth(i) = 0 Then
            DBGrd.Columns(i).Visible = False
        End If
    Next
    'DBGrd.Splits(0).MarqueeStyle=dbgDottedCellBorder
    'DBGrd.Columns(0).Locked = True
    Exit Function
er:
    MsgBox "Format Grid"
End Function
Public Function MoveTo(v As Integer)
    On Error GoTo er
    If Not DataCtrl.Recordset.EOF Then
        DataCtrl.Recordset.Move (v)
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
            'DataCtrl.Recordset.Edit
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
   'DataCtrl.Recordset.Requery
   
    DataCtrl.Refresh
   ' MsgBox DataCtrl.Recordset!batchno
    DBGrd.Col = ColPos
    GoPos

  '  FormatGrd
    Exit Function
er:
    MsgBox "Refresh Error"
End Function
Public Function SortNo()
Dim i As Integer
   ' DataCtrl.Recordset.Resync adAffectAllChapters, adResyncAllValues
   
    If Not DataCtrl.Recordset.EOF() Then
       DataCtrl.Recordset.MoveFirst
       i = 1
       DataCtrl.LockType = adLockOptimistic
       While Not DataCtrl.Recordset.EOF
            DataCtrl.Recordset.Update DataCtrl.Recordset.Fields(0), i
            i = i + 1
            DataCtrl.Recordset.MoveNext
       Wend
    End If
    
    DataCtrl.Refresh
End Function
Public Function Retrieve(ByVal LConn As ADODB.Connection, ByRef Rws As Long, LSvrConStr As String)
    Dim SQL As String
       
       SQL = "INSERT INTO " & TabName _
       & " (" & dFlds & ")  " _
       & " SELECT " & sFlds & " FROM [" _
       & LSvrConStr & "]." & ServerTbName & " WHERE " & fltStr
'MsgBox SQL
        LConn.Execute SQL, Rws
        
End Function
Public Function RsToArray(ByVal iRs As ADODB.Recordset, aVar As Variant, c As Long)
    
    c = iRs.RecordCount
    aVar = iRs.GetRows(c)
End Function
Public Function InsertGrd(ByVal SConn As ADODB.Connection _
, ByVal LConn As ADODB.Connection, ByRef Rws As Long, LSvrConStr As String)
    
  Dim insSQL As String
  Dim delSQL As String
  Dim iConn As New ADODB.Connection
  Dim iCmd As New ADODB.Command
  Dim iRs As New ADODB.Recordset
  Dim iRsx As New ADODB.Recordset
  Dim avarRecords As Variant
  Dim cN As Long
  Dim j, Ins
  Dim cmd As New ADODB.Command
  Dim r As New ADODB.Recordset
    On Error GoTo Qme
    

  
  delSQL = "DELETE FROM " & ServerTbName & " WHERE " & fltStr
    
    SConn.Execute delSQL
    
    ''Ins = "SELECT " & sFlds & " FROM " & TabName

    ''iRsx.Open Ins, LConn, adOpenDynamic, adLockReadOnly

    'Sleep (1000)

    ''RsToArray iRsx, avarRecords, cN

    ''Ins = "SELECT " & dFlds & " FROM " & ServerTbName
    ''iRs.Open Ins, SConn, adOpenKeyset, adLockOptimistic

    'If cN > 0 Then
    '   Rws = cN
    'Else
    '   Rws = 0
     'End If

    Dim t As String
    
    t = "INSERT INTO " & ServerTbName & " (" & dFlds & ") Select " _
        & sFlds & " from [" & LSvrConStr & "]." & TabName
    SConn.Execute t, cN




    ''For cN = 0 To UBound(avarRecords, 2)
      '' iRs.AddNew
       ''For j = 0 To iRs.Fields.Count - 1
         ''  iRs.Fields(j) = avarRecords(j, cN)
           'MsgBox iRs.Fields(j)
      '' Next j
       ''iRs.Update
    ''Next cN

    If cN > 0 Then
       Rws = cN
    Else
       Rws = 0
    End If
'
'Pls. don't remove KKZ  insSQL = "INSERT INTO [" & LSvrConStr & "]." & ServerTbName _
  & " (" & dFlds & ")  " _
  & " SELECT " & sFlds & " FROM " & TabName
       
       
    ''iRs.Close
    ''iRsx.Close
    Exit Function
Qme:
    MsgBox err.Description, vbInformation, "MitAdoGrid.InsertGrd"
End Function


Public Function FillLineNo() As Double
    On Error GoTo er
    
    Dim iRs As New ADODB.Recordset
    Dim iConn As New ADODB.Connection
    Dim i As Integer
    Dim iConString As String
    
    iConString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" _
     & TempDbPath & ";" & TempDbUID & ";" & TempDbPWD
            
     iRs.LockType = adLockOptimistic
     iConn.Open iConString
     iRs.Open "Select [LineNo] from [" & TabName + "] order by [LineNo]", iConn
     
    
    If Not iRs.EOF() Then
       iRs.MoveFirst
    End If
    i = 1
   
    Do While Not iRs.EOF()
       iRs!LineNo = i
       iRs.UpdateBatch
       i = i + 1
       iRs.MoveNext
    Loop
    iRs.Close
    iConn.Close
    Exit Function
er:
End Function
