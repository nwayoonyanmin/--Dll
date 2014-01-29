VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Maintenance"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmUsers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   3240
      Picture         =   "frmUsers.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveFirst 
      Height          =   375
      Left            =   360
      Picture         =   "frmUsers.frx":0168
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveLast 
      Height          =   375
      Left            =   6120
      Picture         =   "frmUsers.frx":0BA4
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdMoveNext 
      Height          =   375
      Left            =   5400
      Picture         =   "frmUsers.frx":1588
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdMovePrevious 
      Height          =   375
      Left            =   1080
      Picture         =   "frmUsers.frx":168C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdNew 
      Height          =   375
      Left            =   1800
      Picture         =   "frmUsers.frx":1790
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   4680
      Picture         =   "frmUsers.frx":1894
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   2520
      Picture         =   "frmUsers.frx":1990
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   375
      Left            =   3960
      Picture         =   "frmUsers.frx":1A94
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdBrowseEmployeeID 
      Height          =   375
      Left            =   3240
      Picture         =   "frmUsers.frx":1F58
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   375
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Left            =   4200
      Picture         =   "frmUsers.frx":2574
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   375
      Left            =   3720
      Picture         =   "frmUsers.frx":2B90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.ComboBox cboLevel 
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtEmployeeID 
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtRemark 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox txtFullName 
      Height          =   285
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtLevel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ActiveConnection As ADODB.Connection
Public UseActiveConnection As Boolean

Public Greet As Boolean
Public ConnString As String
Public ConnType As Byte
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

Dim Dat1 As New MitUniversalClass
Dim MyGeneral As New MitGeneral
Private Sub cmdBrowseEmployeeID_Click()
    MsgBox "Not Available"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub


Private Sub cmdNew_Click()
    ClearScreen
End Sub

Private Sub cmdPrint_Click()
    MsgBox "Not Available"
End Sub
Private Sub Form_Load()
    Dat1.ConString = ConnString
    Dat1.ConnectionType = ConnType
    If UseActiveConnection Then
        Dat1.UseActiveConnection = True
        Set Dat1.ActiveConnection = ActiveConnection
    End If
    Select Case ProjectType
        Case 3
            Dat1.TabName = "Teller"
            Dat1.NoOfText = 5
            Dat1.NoOfNumeric = 1
            Dat1.OrderBy = "tID"
            Dat1.Initialize
            Dat1.DefaultFieldNames
            
            Dat1.LetTextField 1, "tID"
            Dat1.LetTextField 2, "tPassword"
            Dat1.LetTextField 3, "tName"
            Dat1.LetTextField 4, "Post"
            Dat1.LetTextField 5, "Remark"
            Dat1.LetNumericField 1, "trights"
            Dat1.LoadToCombo cboLevel, "UserLevels", "UserLevel", "Description"
        Case Else
            cmdBrowse.Visible = True
            Dat1.TabName = "Users"
            Dat1.NoOfText = 6
            Dat1.NoOfNumeric = 1
            Dat1.OrderBy = "UserName"
            Dat1.Initialize
            Dat1.DefaultFieldNames
            
            Dat1.LetTextField 1, "UserName"
            Dat1.LetTextField 2, "UserPassword"
            Dat1.LetTextField 3, "FullName"
            Dat1.LetTextField 4, "Description"
            Dat1.LetTextField 5, "Remark"
            Dat1.LetTextField 6, "EmployeeID"
            Dat1.LetNumericField 1, "UserLevel"

            Dat1.LoadToCombo cboLevel, "UserLevels", "UserLevel", "Description"
    End Select
    
    ClearScreen
    If usrLevel < 100 Then
        cmdMoveFirst.Enabled = False
        cmdMovePrevious.Enabled = False
        cmdMoveNext.Enabled = False
        cmdMoveLast.Enabled = False
        cmdDelete.Enabled = False
        cmdNew.Enabled = False
        cmdSearch.Enabled = False
        cmdBrowse.Enabled = False
        txtName.Enabled = False
        cboLevel.Enabled = False
        
        txtName = usrName
        txtPassword = MyGeneral.Encrypt(usrPassword, EncryptionCode)
        cmdSearch_Click
    End If
End Sub
Sub UpdateObject()
    Dat1.LetText 1, txtName
 
        Dat1.LetText 2, MyGeneral.Encrypt(txtPassword, EncryptionCode)

    Dat1.LetNumeric 1, Val(txtLevel)
    Select Case ProjectType
        Case 3
            Dat1.LetText 3, txtFullName
            Dat1.LetText 4, txtDescription
            Dat1.LetText 5, txtRemark
        Case Else
            Dat1.LetText 3, txtFullName
            Dat1.LetText 4, txtDescription
            Dat1.LetText 5, txtRemark
            Dat1.LetText 6, txtEmployeeID
    End Select
    
End Sub
Sub ClearScreen()
    txtName = ""
    txtPassword = ""
    cboLevel.ListIndex = 0
    txtFullName = ""
    txtDescription = ""
    txtRemark = ""
    txtEmployeeID = ""
End Sub
Sub UpdateScreen()
    txtName = Dat1.GetText(1)
        txtPassword = MyGeneral.Decrypt(Dat1.GetText(2), EncryptionCode)
    Select Case ProjectType
        Case 3
            txtFullName = Dat1.GetText(3)
            txtDescription = Dat1.GetText(4)
            txtRemark = Dat1.GetText(5)
            'txtEmployeeID = dat1.GetText(6)
        Case Else
            txtFullName = Dat1.GetText(3)
            txtDescription = Dat1.GetText(4)
            txtRemark = Dat1.GetText(5)
            txtEmployeeID = Dat1.GetText(6)
    End Select
    txtLevel = Dat1.GetNumeric(1)
End Sub
Private Sub cmdSearch_Click()
    Select Case ProjectType
        Case 3
            Dat1.SearchWhere = "tID='" + txtName + "'"
        Case Else
            Dat1.SearchWhere = "UserName='" + txtName + "'"
    End Select
    If Dat1.ReadIt = 255 Then
        UpdateScreen
    Else
        MsgBox "Record Not Found"
        ClearScreen
        Beep
    End If
End Sub

Private Sub txtLevel_Change()
        For i = 0 To cboLevel.ListCount - 1
        If cboLevel.ItemData(i) = Val(txtLevel) Then
            cboLevel.ListIndex = i
            Exit Sub
        End If
    Next
End Sub
Private Sub cboLevel_Click()

    txtLevel = cboLevel.ItemData(cboLevel.ListIndex)
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch_Click
    End If
End Sub
Private Sub cmdBrowse_Click()
    Dim t As String
    
    If ConnType Mod 100 = 1 Then
        t = "A"
    Else
        t = ""
    End If
    a = Dat1.Browse(t, "", _
                 " Select UserName,FullName from Users", "UserName", _
                 " Users Browsers", 8000, 3000, "User's Name,User's Full Name", "2000,3500")
    If a <> "" Then
        txtName = a
        cmdSearch_Click
    End If
End Sub
Private Sub cmdMoveNext_Click()
    Select Case ProjectType
        Case 3
            Dat1.SearchWhere = "tID>='" + txtName + "'"
        Case Else
            Dat1.SearchWhere = "UserName>='" + txtName + "'"
    End Select
    If Dat1.MoveNext = 255 Then
        UpdateScreen
    End If
End Sub

Private Sub cmdMovePrevious_Click()
    Select Case ProjectType
        Case 3
            Dat1.SearchWhere = "tID<='" + txtName + "'"
        Case Else
            Dat1.SearchWhere = "UserName<='" + txtName + "'"
    End Select
    If Dat1.MovePrevious = 255 Then
        UpdateScreen
    End If
End Sub
Private Sub cmdMoveFirst_Click()
    If Dat1.MoveFirst = 255 Then
        UpdateScreen
    End If
End Sub
Private Sub cmdMoveLast_Click()
    If Dat1.MoveLast = 255 Then
        UpdateScreen
    End If
End Sub
Private Sub cmdSave_Click()
    UpdateObject
    Select Case ProjectType
        Case 3
            Dat1.SearchWhere = " tID='" + txtName + "' "        'Key Information
        Case Else
            Dat1.SearchWhere = " UserName='" + txtName + "' "        'Key Information
    End Select
    If Dat1.UpdateSQL = 255 Then
        MsgBox "Updated Successfully"
    Else
        If Dat1.InsertSQL = 255 Then
            MsgBox "New Record Saved Successfully"
        Else
            MsgBox "Can not be saved nor inserted"
        End If
    End If
End Sub
Private Sub cmdDelete_Click()
    Select Case ProjectType
        Case 3
        Case Else
            Dat1.SearchWhere = "UserName='" + Text1 + "'"         'Key Information
    End Select
    If Dat1.DeleteIt = 255 Then
        MsgBox "Deleted Successfully"
    Else
        MsgBox "Can not be deleted"
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        On Error Resume Next
    If KeyCode = 116 Then Exit Sub
    Select Case KeyCode
        Case 27
               Unload Me
    End Select
End Sub
Public Function EditUsers()
    frmUsers.Show 1
End Function
