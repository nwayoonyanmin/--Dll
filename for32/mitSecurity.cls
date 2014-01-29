VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mitSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Greet As Boolean
Public ProjectType As Byte
Public Encrypted As Boolean
Public ActiveConnection As ADODB.Connection
Public UseActiveConnection As Boolean
Public ConnString As String
Public ConnType As Byte


Public Level As Long
Public Name As String
Public PassWord As String
Public FullName As String
Public Description As String
Public Remark As String
Public EmployeeID As String
Public EncryptionCode As Byte
Public BioMetricPath_S As String
Public UseBioMetric As Boolean
Public UseSmartCard As Boolean
Public Function Security() As Long
    frmSecurity.ProjectType = ProjectType
    frmSecurity.Greet = Greet
    frmSecurity.ConnString = ConnString
    frmSecurity.EncryptionCode = EncryptionCode
    frmSecurity.BioMetricPath = BioMetricPath_S
    frmSecurity.UseBioMetric_L = UseBioMetric
    frmSecurity.UseSmartCard_L = UseSmartCard
    If UseActiveConnection Then
        frmSecurity.UseActiveConnection = True
        Set frmSecurity.ActiveConnection = ActiveConnection
    End If
    Security = frmSecurity.Security
    If Security > 0 Then
        Name = frmSecurity.usrName
        Level = frmSecurity.usrLevel
        PassWord = frmSecurity.usrPassword
        FullName = frmSecurity.usrFullName
    End If
End Function
Public Function EditUsers() As Long
    frmUsers.usrName = Name
    frmUsers.usrPassword = PassWord
    frmUsers.ProjectType = ProjectType
    frmUsers.ConnString = ConnString
    frmUsers.ConnType = ConnType
    frmUsers.usrLevel = Level
    frmUsers.usrName = Name
    frmUsers.EncryptionCode = EncryptionCode
    If UseActiveConnection Then
        frmUsers.UseActiveConnection = True
        Set frmUsers.ActiveConnection = ActiveConnection
    End If
    frmUsers.EditUsers
End Function
