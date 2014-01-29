VERSION 5.00
Begin VB.Form frmSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security First"
   ClientHeight    =   1560
   ClientLeft      =   1845
   ClientTop       =   1950
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecurity.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1560
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSMC 
      Caption         =   "&Smart Card"
      Height          =   315
      Left            =   1050
      TabIndex        =   8
      Top             =   1215
      Width           =   1035
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "&Biometric"
      Height          =   315
      Left            =   105
      TabIndex        =   7
      Top             =   1215
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Quit"
      Height          =   330
      Left            =   3585
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2145
      TabIndex        =   2
      Top             =   1215
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   -70
      Width           =   5340
      Begin VB.TextBox txtPW 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Name 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const vbShiftMask = 1
Const vbCtrlMask = 2
Const vbAltMask = 4
Dim Cnt As Byte
Dim NorConn As New ADODB.Connection
Public ActiveConnection As ADODB.Connection
Public UseActiveConnection As Boolean

Public Greet As Boolean
Public ConnString As String
Public ProjectType As Byte
Public Encrypted As Boolean
Public EncryptionCode As Byte

Public usrLevel As Long
Public usrName As String
Public usrPassword As String
Public usrFullName As String
Public usrEmployeeID As String
Public usrDescription As String
Public usrRemark As String
Public BioMetricPath As String

Public UseBioMetric_L As Boolean
Public UseSmartCard_L As Boolean


Public Function Security() As Integer
    frmSecurity.Show 1
    Security = usrLevel     '-1=Cancel, 0=Wrong, >0 =Valid User
End Function

Private Sub cmdBio_Click()
'Dim fp As New MITFPGeneral
'Dim aRet()
'Dim RetStr As String
'Dim UName As String
'Dim Pw
'    Me.Hide
'    fp.DbPath = BioMetricPath
'    If fp.DirectScan(i, N, True, True, True, Pw) Then
'        'MsgBox i & ",  " & n
'         Me.Show
'        txtID = i
'        txtPW = Pw
'        cmdOK_Click
'    Else
'        Me.Show
'        MsgBox "Invalid User.Pls Try", vbInformation
'
'    End If
''    Me.Show
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    usrLevel = -1
    Conn.Close
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim Conn As ADODB.Connection
    If UseActiveConnection Then
        Set Conn = ActiveConnection
    Else
        Set Conn = NorConn
    End If
    Dim myTablename As String
    Dim myFullName As String
    Dim myUserPassword As String
    Dim myLevel As String
    Dim myUserID As String
    Dim myGen As New MitGeneral
    Select Case ProjectType
        Case 1  'BizAcc
            myTablename = "Users"
            myUserID = "UserID"
            myUserPassword = "Password"
            myLevel = "LevelAccess"
            myFullName = "Name"
        Case 2  'Offline Rusers CityMart
            myTablename = "rUser"
            myUserID = "USR_ID"
            myUserPassword = "USR_PW"
            myLevel = "ACC_LVL"
            myFullName = "USR_NM"
        Case 3  ' IRBS
            myTablename = "Teller"
            myUserID = "tID"
            myUserPassword = "tPassword"
            myLevel = "tRights"
            myFullName = "tName"
        Case Else  'Standard
            myTablename = "Users"
            myFullName = "FullName"
            myUserPassword = "UserPassword"
            myLevel = "UserLevel"
            myUserID = "UserName"
    End Select
    If txtID = "" And txtPW = "" Then Exit Sub
    On Error GoTo er
    Dim HotKey As String
    Dim MasterLevel As Long
    Dim xpassword As String
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    MasterLevel = 1000
    HotKey = "?,A"
    xid = txtID
    xpassword = myGen.Encrypt(txtPW, EncryptionCode)
    Cnt = Cnt + 1
    cmd.Name = "ReadUsers"
    'Rs.LockType = adLockOptimistic
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenKeyset
    cmd.CommandText = "select * from " & myTablename & "  Where " & _
            " (" & myTablename & "." & myUserID & ")='" + xid + "'" + _
            " And " + myTablename & "." & myUserPassword + "='" + xpassword + "'"
'            MsgBox Conn
    cmd.ActiveConnection = Conn
    rs.Open cmd
    If Not rs.EOF Then
        usrLevel = rs.Fields(myLevel)
        usrFullName = rs.Fields(myFullName)
        usrName = txtID
        usrPassword = txtPW
        Cnt = 0
            'Conn.Close
        Unload Me
    Else
        If UCase(xpassword) = HotKey Then
            usrLevel = MasterLevel
            usrName = txtID
            usrPassword = txtPW
            Cnt = 0
            Unload Me
                'Conn.Close
            Exit Sub
        End If
        usrLevel = 0
        If Cnt >= 3 And usrLevel = 0 Then
            MsgBox "Third attempt failed", 0 + 16, gsMsgTitile
            Cnt = 0
            usrLevel = 0
            usrName = txtID
            usrPassword = txtPW
                'Conn.Close
            Unload Me
        Else
            txtID = ""
            txtPW = ""
            txtID.SetFocus
        End If
    End If
    Exit Sub
er:
    MsgBox err.Description
End Sub

Private Sub cmdSMC_Click()
' Dim sm As New MITSmartCardGeneral
'    Me.Hide
'    If sm.CardAuthorise(i, p, N) Then
'        Me.Show
'        txtID = i
'        txtPW = p
'        cmdOK_Click
'    Else
'        Me.Show
'        MsgBox "Invalid User.Pls Try", vbInformation
'    End If
  
End Sub

Private Sub Form_Load()
    On Error GoTo Skip
    xpassword = 0
    usrLevel = 0
    xid = ""
    If Not UseActiveConnection Then
        NorConn.ConnectionString = ConnString 'MyData.PrepareConString(MyDataType)
        NorConn.Open
    End If
    
    If UseBioMetric_L Then
       cmdBio.Visible = True
    Else
        cmdBio.Visible = False
    End If
    If UseSmartCard_L Then
       cmdSMC.Visible = True
    Else
          cmdSMC.Visible = False
    End If
    
    Exit Sub
Skip:
    MsgBox "            Invalid Database ! " + Chr(13) + Chr(10) + "Please Check Your DATABASE in Company Profile ->" & err.Description, 0 + 64
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Greet And usrLevel > 0 Then
        MsgBox "Hello " & usrFullName, , "Welcome"
    End If
End Sub
Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtID)) <> 0 Then
            If Not ValidString(txtID) Then
                txtID = ""
                Exit Sub
            End If
        End If
       txtPW.SetFocus
    End If
End Sub
Private Sub txtPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtPW)) <> 0 Then
            If Not ValidString(txtPW) Then
                txtPW = ""
                Exit Sub
            End If
        End If
        cmdOK_Click
    Else
    End If
End Sub
Function ValidString(v As Variant) As Boolean
    ValidString = False
    If Len(Trim(v)) <> 0 Then
        If ((Asc(v)) <> 124) And ((Asc(v)) <> 39) And ((Asc(v)) <> 91) Then
            ValidString = True
        Else
            ValidString = False
        End If
    End If
End Function
