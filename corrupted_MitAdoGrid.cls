8    �  Y                    N          y y            Y        	          
Y      	       gj�
Y     	         �
Y     	       `�
Y     	       �q�Y     	         Y     	         Y     	         Y     	         Y  	  	 	          
Y  
  
 	       ��Y     	       (  Y     	       0   L i n e N o  S t o c k I d  D e s c r i p t i o n  L o c a t i o n  C u r C o d e  C u r R a t e 
 P r i c e  Q t y  C o s t  A m o u n t  o C u r C o d e  R g s A m t  S e r v i c e N o �    ����������������������y z  ���     Y           ����         P r i m a r y K e y ��                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         �v1                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         �x                       @�                                                                                                                                                                                                                                                                                                                                                                                                                                                             �   { �   {�   {�   {�   {�   {�   {                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 Rx       �D��*�n                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            �?                                       0 0 0 0 0 0 U S $ U S $ N H B B B 6  �          �?                                       0 0 0 0 0 0 U S $ U S $ N H B B B 6  �          �?                                       0 0 0 0 0 0 U S $ U S $ N H B B B 6  �          �?                                       0 0 0 0 0 0 U S $ U S $ N H B B B 6  �          �?                                       0 0 0 0 0 0 U S $ U S $ N H B B B 6  �          �?                                       0 0 0 0 0 0 U S $ U S $ N H B B B 6  �          �?                                       0 0 0 0 0 0 U S $ U S $ N H B B B 6  �VERSION 1.0 CLASS
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
      