VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form bwrUniversal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Universal Browser"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   Icon            =   "bwrUniversal.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "bwrUniversal.frx":030A
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "bwrUniversal.frx":031E
      TabIndex        =   0
      Top             =   1080
      Width           =   9225
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2445
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "bwrUniversal.frx":0CE9
      Left            =   2625
      List            =   "bwrUniversal.frx":0CFF
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2445
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Width           =   1250
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2625
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   315
      Left            =   4500
      Picture         =   "bwrUniversal.frx":0D55
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Refresh Card Limit"
      Top             =   120
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6360
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblCount 
      Caption         =   "Count"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "bwrUniversal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ComboFields() As String
Dim ComboFieldTypes() As String

Private Sub cmdRefresh_Click()
    DoQuery
End Sub

Private Sub Combo2_Click()
    If Combo2.ListIndex = 5 Then
        Text2.Visible = True
    Else
        Text2.Visible = False
    End If
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    Select Case V6 ' Width
        Case 0
        Case Else
            Me.Width = V6
            DBGrid1.Width = V6 - 300
            'Text1.Width = V6 - 5000
            cmdSelect.Left = V6 / 2 - 1300
            cmdCancel.Left = V6 / 2 + 200
    End Select
    
    Select Case V7 ' Width
        Case 0
        Case Else
            diff = V7 - Me.Height
            Me.Height = V7
            cmdSelect.Top = cmdSelect.Top + diff
            cmdCancel.Top = cmdCancel.Top + diff
            Data1.Top = Data1.Top + diff
            lblCount.Top = lblCount.Top + diff
    End Select
    If Combo1.ListCount > 0 Then
        DBGrid1.Height = Me.Height - 1900 - 250
    Else
        DBGrid1.Height = Me.Height - 1200
    End If
    Me.Caption = V5
    If UCase(Left(V1, 1)) = "A" Then
        If V2 <> "" Then
            Data1.Connect = V2      ' Connection String In Case of Password
        End If
        If v100 = "" Then
            Data1.DatabaseName = DataPath(V2)   ' Database Path
        Else
            Data1.DatabaseName = v100   ' Database Path
        End If
    Else
        Data1.Connect = V2          ' Connection String for ODBC
    End If
    If UCase(Mid(V1, 2, 1)) = "F" Then      'AM12Times
        DBGrid1.Font.Size = Val(Mid(V1, 3, 2))
        DBGrid1.HeadFont.Size = Val(Mid(V1, 3, 2))
        Combo1.Font.Size = Val(Mid(V1, 3, 2))
        Text1.Font.Size = Val(Mid(V1, 3, 2))
        Text2.Font.Size = Val(Mid(V1, 3, 2))
        If Mid(V1, 5, 30) <> 0 Then
            DBGrid1.Font.Name = Mid(V1, 5, 30)
            DBGrid1.HeadFont.Name = Mid(V1, 5, 30)
            Combo1.Font.Name = Mid(V1, 5, 30)
            Text1.Font.Name = Mid(V1, 5, 30)
            Text2.Font.Name = Mid(V1, 5, 30)
        End If
    Else
        s = DBGrid1.Font
    End If
    
    If v14 = "" Then
        Data1.RecordSource = V3
    Else
        Data1.RecordSource = V3 & " where " & v14
    End If
    
    If v15 <> "" Then
        Data1.RecordSource = Data1.RecordSource & " order by " & v15
    End If
    
    Data1.Refresh
    FormatGrid
    If Combo1.ListCount = 0 Then
        DoCombo
    End If
End Sub
Sub FormatGrid()
    Dim a() As Variant
    Dim N As Integer
    'DBGrid1.Columns = N
    If V8 = "" Then ' Column Heading
    Else
        N = BreakUp(V8, a())
        For i = 1 To N
            DBGrid1.Columns(i - 1).Caption = a(i - 1)
        Next
    End If
    If V9 = "" Then ' Column Width
    Else
        N = BreakUp(V9, a())
        For i = 1 To N
            If Val(a(i - 1)) <> 0 Then
                DBGrid1.Columns(i - 1).Width = Val(a(i - 1))
            End If
        Next
    End If
    If V10 = "" Then ' Column Format
    Else
        N = BreakUp(V10, a())
        For i = 1 To N
            If Trim(a(i - 1)) <> "" Then
                DBGrid1.Columns(i - 1).NumberFormat = a(i - 1)
            End If
        Next
    End If
End Sub
Sub DoCombo()
    Dim a() As Variant
    Dim N As Integer
    If V11 = "" Then ' Combo List
        DBGrid1.Top = 0
    Else
        DBGrid1.Height = Me.Height - 2500
        N = BreakUp(V11, a())
        For i = 1 To N
            Combo1.AddItem a(i - 1)
        Next
        Combo1.ListIndex = 0
        Combo2.ListIndex = 1
        N = BreakUp(V12, a()) ' Combo Fileds
        ReDim ComboFields(N)
        For i = 1 To N
            ComboFields(i - 1) = a(i - 1)
        Next
        
        N = BreakUp(V13, a()) ' Combo FiledTypes
        ReDim ComboFieldTypes(N)
        For i = 1 To N
            ComboFieldTypes(i - 1) = a(i - 1)
        Next
    End If
End Sub
Sub DoQuery()
    On Error GoTo DoSomething
    If Text1 = "" Then
        Form_Load
        Exit Sub
    End If
    Select Case UCase(Trim(LTrim(ComboFieldTypes(Combo1.ListIndex))))
    Case "T"
        Select Case Combo2.ListIndex
            Case 0  'Exact
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " = " & "'" & Text1 & "'"
            Case 1  'Begins
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " like " & "'" & Text1 & "*'"
            Case 2  'Contains
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " like " & "'*" & Text1 & "*'"
            Case 3  'Less
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " <= " & "'" & Text1 & "'"
            Case 4  'Greater
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " >= " & "'" & Text1 & "'"
        End Select
    Case "N"
        Select Case Combo2.ListIndex
            Case 0  'Exact
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " = " & "" & Text1 & ""
            Case 3  'Less
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " <= " & "" & Text1 & ""
            Case 4  'Greater
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " >= " & "" & Text1 & ""
                d2 = Right(Text1, 10)
            Case 5 'between
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " >= " & " " & Text1 & " and " & _
                        ComboFields(Combo1.ListIndex) & _
                        " <= " & " " & Text2 & " "
        End Select
    Case "D"
        Select Case Combo2.ListIndex
            Case 0  'Exact
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " = " & "#" & Format(Left(Text1, 10), "mm/dd/yyyy") & "#"
            Case 3  'Less
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " <= " & "#" & Format(Left(Text1, 10), "mm/dd/yyyy") & "#"
            Case 4  'Greater
                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " >= " & "#" & Format(Left(Text1, 10), "mm/dd/yyyy") & "#"
            Case 5 'Between

                Data1.RecordSource = V3 + " Where " & ComboFields(Combo1.ListIndex) & _
                        " >= " & "#" & Format(Text1, "mm/dd/yyyy") & "# and " & _
                        ComboFields(Combo1.ListIndex) & _
                        " <= " & "#" & Format(Text2, "mm/dd/yyyy") & "# "
                        
        End Select
    End Select
    If v14 <> "" Then
        Data1.RecordSource = Data1.RecordSource & " and " & v14
    End If
    If v15 <> "" Then
        Data1.RecordSource = Data1.RecordSource & " order by " & v15
    End If
    If b0 Then 'do debug
        MsgBox Data1.RecordSource
    End If
    Data1.Refresh
    FormatGrid

    Exit Sub
DoSomething:
    MsgBox "Your query can not be processed. " & err.Description
End Sub

Private Sub cmdCancel_Click()
    vReturn = ""
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error Resume Next
    vReturn = Data1.Recordset.Fields(V4)
    Unload Me
End Sub

Private Sub Image1_Click()
    DBGrid1.AllowUpdate = True
End Sub

Private Sub lblCount_Click()
    Data1.Recordset.MoveLast
    lblCount = Data1.Recordset.RecordCount
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DoQuery
    End If
End Sub

Function DataPath(V2 As Variant) As String
    Dim g As New MitGeneral
    Dim p() As Variant
    If g.BreakUp(V2, p(), ";") > 0 Then
        DataPath = Mid(p(1), 13, 30)
    End If
End Function
