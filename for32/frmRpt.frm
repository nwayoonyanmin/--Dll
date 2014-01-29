VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmRpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MIT Printing Control Panel"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport cyst 
      Left            =   3360
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Timer Timer1 
      Left            =   90
      Top             =   1905
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Exit"
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
      Left            =   2325
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   1005
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   165
      TabIndex        =   5
      Top             =   0
      Width           =   3855
      Begin VB.OptionButton Option3 
         Caption         =   "Write to a file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Print to printer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Display on screen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    On Error GoTo er

            cyst.ReportFileName = V2
            If Option1.Value = True Then
                cyst.Destination = crptToWindow
            Else
                If Option2.Value = True Then
                    cyst.Destination = crptToPrinter
                Else
                    cyst.Destination = crptToFile
                End If
            End If
            cyst.WindowState = crptMaximized
            
            cyst.Action = 1
            
            For i = 1 To N
                cyst.Formulas(i - 1) = ""
            Next
    Exit Sub
er:
    MsgBox "Can perform printing " & err.Description
End Sub
Private Sub Form_Load()
    
    Dim f() As Variant
    N = BreakUp(V4, f())
    For i = 0 To 100
        cyst.Formulas(i) = ""
    Next
    For i = 1 To N
        cyst.Formulas(i - 1) = f(i - 1)
    Next
    cyst.DataFiles(0) = V1 'MDB
    cyst.ReportFileName = V2 'RPT
    If V6 Then 'PrinterDirect
        cmdOK.Enabled = False
        cyst.Destination = crptToPrinter
        cyst.Action = 1
        Timer1.Interval = 100
    End If
End Sub
Private Sub Timer1_Timer()
    Unload Me
End Sub

