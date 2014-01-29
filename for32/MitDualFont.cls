VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MitDualFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public NormalFont As String
Public NormalFontSize As Byte
Public NormalHeight As Integer


Public Font As String
Public FontSize As Byte
Public Height As Integer
Public SpecialCode As String
Public Sub UpdateScreen(Optional t, Optional Chk)
    If Left(t, 1) = SpecialCode Then
        ChangeFont t
        Chk.Value = 1
    Else
        UnChangeFont t
        Chk.Value = 0
    End If
End Sub
Function UpdateObject(Optional t) As String
    If t.Font = Font Then
        UpdateObject = SpecialCode + t.Text
    Else
        UpdateObject = t.Text
    End If
End Function
Sub UnChangeFont(Optional t)
        t.Font = NormalFont
        t.FontSize = NormalFontSize
        t.Height = IIf(NormalHeight = 0, t.Height, NormalHeight)
End Sub

Sub ChangeFont(Optional t)
        t.Font = Font
        t.FontSize = FontSize
        If Mid(t.Text, 1, 1) = SpecialCode Then
            t.Text = Mid(t.Text, 2, 256)
        End If
        t.Height = IIf(Height = 0, t.Height, Height)
End Sub

