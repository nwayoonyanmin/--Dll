Attribute VB_Name = "modMitDll1"
Public V1 As Variant
Public V2 As Variant
Public V3 As Variant
Public V4 As Variant
Public V5 As Variant
Public V6 As Variant
Public V7 As Variant
Public V8 As Variant
Public V9 As Variant
Public V10 As Variant
Public V11 As Variant
Public V12 As Variant
Public V13 As Variant
Public v14 As Variant
Public v15 As Variant
Public v16 As Variant

Public v100 As Variant

Public vReturn As Variant
Public b0 As Boolean
Public ColHeadings As String


Function DoBrowse(DataType As String, ConnectionOrFile As String, _
    SqlString As String, ReturnKeys As String, _
    Optional FormName, Optional FormWidth, Optional FormHeight, _
    Optional Colheading, Optional ColWidth, Optional ColFormat, _
    Optional ComboList, Optional ComboFields, Optional ComboFieldTypes, _
    Optional DefaultWhere, Optional OrderBy, Optional InitialPosition, _
    Optional DoDebug, Optional DatabaseName) As Variant
    
    b0 = DoDebug
    V1 = DataType
    V2 = ConnectionOrFile
    V3 = SqlString
    V4 = ReturnKeys
    V5 = FormName
    V6 = IIf(IsMissing(FormWidth), 0, FormWidth)
    V7 = IIf(IsMissing(FormHeight), 0, FormHeight)
    V8 = IIf(IsMissing(Colheading), "", Colheading)
    V9 = IIf(IsMissing(ColWidth), "", ColWidth)
    V10 = IIf(IsMissing(ColFormat), "", ColFormat)
    V11 = IIf(IsMissing(ComboList), "", ComboList)
    V12 = IIf(IsMissing(ComboFields), "", ComboFields)
    V13 = IIf(IsMissing(ComboFieldTypes), "", ComboFieldTypes)
    v14 = IIf(IsMissing(DefaultWhere), "", DefaultWhere)
    v15 = IIf(IsMissing(OrderBy), "", OrderBy)
    v16 = IIf(IsMissing(InitialPosition), "", InitialPosition)
    v100 = IIf(IsMissing(DatabaseName), "", DatabaseName)
    If DoDebug Then
        MsgBox "V2: " & V2 & "  V100: " & v100
    End If
    bwrUniversal.Show 1
    'DoBrowse = WorkspaceTypeEnum.dbUseODBC
    DoBrowse = vReturn
End Function
Function DoReport(MDB As String, RPT As String, _
                   DefaultDestination As Byte, _
                  Optional Formula, _
                  Optional DefaultWindowState, _
                  Optional PrinterDirect)
    V1 = MDB
    V2 = RPT
    V3 = DefaultDestination
    V4 = IIf(IsMissing(Formula), 0, Formula)
    V5 = IIf(IsMissing(DefaultWindowState), 0, DefaultWindowState)
    V6 = IIf(IsMissing(PrinterDirect), False, PrinterDirect)
    frmRpt.Show 1
End Function
Function BreakUp(s As Variant, a() As Variant, Optional Sep) As Integer
    If IsMissing(Sep) Then
        Sep = ","
    End If
    Dim StrLen As Integer
    Dim Temp As Variant
    StrLen = Len(s)
    For i = 1 To StrLen
        
        If Mid(s, i, 1) = Sep Or i = StrLen Then
            BreakUp = BreakUp + 1
            ReDim Preserve a(BreakUp)
            a(BreakUp - 1) = Temp
            If i = StrLen Then
                 a(BreakUp - 1) = a(BreakUp - 1) + Mid(s, i, 1)
            End If
            Temp = ""
        Else
            Temp = Temp + Mid(s, i, 1)
        End If
    Next
    
End Function

