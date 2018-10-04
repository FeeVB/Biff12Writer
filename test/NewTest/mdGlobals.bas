Attribute VB_Name = "mdGlobals"
' A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets
Option Explicit
DefObj A-Z

'获得活动窗口句柄的API
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Const MODULE_NAME As String = "mdGlobals"

Private Const GMEM_DDESHARE As Long = &H2000
Private Const GMEM_MOVEABLE As Long = &H2

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long


Private Const VT_I8 As Long = &H14
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Function ApiEmptyByteArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal VarType As VbVarType = vbByte, Optional ByVal Low As Long = 0, Optional ByVal Count As Long = 0) As Byte()
Private Declare Function ApiDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VariantChangeType Lib "oleaut32" (Dest As Variant, src As Variant, ByVal wFlags As Integer, ByVal vt As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function ApiCreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long



'mdBiff12Shared
#Const ImplUseShared = BIFF12_USESHARED
#Const ImplPublicClasses = BIFF12_PUBLICCLASSES

#If ImplPublicClasses = 0 Then

Public Type UcsBiff12BrtColorType
    m_xColorType        As Byte
    m_index             As Byte
    m_nTintAndShade     As Integer
    m_bRed              As Byte
    m_bGreen            As Byte
    m_bBlue             As Byte
    m_bAlpha            As Byte
End Type

Public Type UcsBiff12BrtFontType
    m_dyHeight          As Integer
    m_grbit             As Integer
    m_bls               As Integer
    m_sss               As Integer
    m_uls               As Byte
    m_bFamily           As Byte
    m_bCharSet          As Byte
    '--- padding 1 bytes
    m_brtColor          As UcsBiff12BrtColorType
    m_bFontScheme       As Byte
    '--- padding 3 bytes
    m_name              As String
End Type

Public Type UcsGradientStopType
    brtColor            As UcsBiff12BrtColorType
    xnumPosition        As Double
End Type

Public Type UcsBiff12BrtFillType
    m_fls               As Long
    m_brtColorFore      As UcsBiff12BrtColorType
    m_brtColorBack      As UcsBiff12BrtColorType
    m_iGradientType     As Long
    m_xnumDegree        As Double
    m_xnumFillToLeft    As Double
    m_xnumFillToRight   As Double
    m_xnumFillToTop     As Double
    m_xnumFillToBottom  As Double
    m_cNumStop          As Long
    m_xfillGradientStop() As UcsGradientStopType
End Type

Public Type UcsBiff12BrtBlxfType
    m_dg                As Integer
    '--- padding 2 bytes
    m_brtColor          As UcsBiff12BrtColorType
End Type

Public Type UcsBiff12BrtBorderType
    m_flags             As Long
    m_blxfTop           As UcsBiff12BrtBlxfType
    m_blxfBottom        As UcsBiff12BrtBlxfType
    m_blxfLeft          As UcsBiff12BrtBlxfType
    m_blxfRight         As UcsBiff12BrtBlxfType
    m_blxfDiag          As UcsBiff12BrtBlxfType
End Type

Public Type UcsBiff12BrtXfType
    m_ixfeParent        As Integer
    m_iFmt              As Integer
    m_iFont             As Integer
    m_iFill             As Integer
    m_ixBorder          As Integer
    m_trot              As Byte
    m_indent            As Byte
    m_flags             As Integer
    m_xfGrbitAtr        As Byte
End Type

Public Type UcsBiff12BrtStyleType
    m_ixf               As Long
    m_grbitObj1         As Integer
    m_iStyBuiltIn       As Byte
    m_iLevel            As Byte
    m_stName            As String
End Type

Public Type UcsBiff12BrtWbPropType
    m_flags             As Long
    m_dwThemeVersion    As Long
    m_strName           As String
End Type

Public Type UcsBiff12BrtBookViewType
    m_xWn               As Long
    m_yWn               As Long
    m_dxWn              As Long
    m_dyWn              As Long
    m_iTabRatio         As Long
    m_itabFirst         As Long
    m_itabCur           As Long
    m_flags             As Integer
End Type

Public Type UcsBiff12BrtBundleShType
    m_hsState           As Long
    m_iTabID            As Long
    m_strRelID          As String
    m_strName           As String
End Type

Public Type UcsBiff12BrtWsPropType
    m_flags             As Long
    m_brtcolorTab       As UcsBiff12BrtColorType
    m_rwSync            As Long
    m_colSync           As Long
    m_strName           As String
End Type

Public Type UcsBiff12BrtColInfoType
    m_colFirst          As Long
    m_colLast           As Long
    m_colDx             As Long
    m_ixfe              As Long
    m_flags             As Integer
End Type

Public Type UcsBiff12BrtColSpanType
    m_colMic            As Long
    m_colLast           As Long
End Type

Public Type UcsBiff12BrtRowHdrType
    m_rw                As Long
    m_ixfe              As Long
    m_miyRw             As Integer
    '--- padding
    m_flags             As Long '-- 3 bytes
    m_ccolspan          As Long
    m_rgBrtColspan()    As UcsBiff12BrtColSpanType
End Type

Public Type UcsBiff12BrtFmtType
    m_iFmt              As Integer
    m_stFmtCode         As String
End Type

Public Type UcsBiff12UncheckedRfXType
    m_rwFirst           As Long
    m_rwLast            As Long
    m_colFirst          As Long
    m_colLast           As Long
End Type

Public Type UcsBiff12BrtFileVersionType
    m_guidCodeName      As String
    m_stAppName         As String
    m_stLastEdited      As String
    m_stLowestEdited    As String
    m_stRupBuild        As String
End Type

#End If ' ImplPrivateClasses

#If ImplUseShared = 0 Then

Public Enum UcsHorAlignmentEnum
    ucsHalLeft = 0
    ucsHalRight = 1
    ucsHalCenter = 2
End Enum

Public Enum UcsVertAlignmentEnum
    ucsValTop = 0
    ucsValMiddle = 1
    ucsValBottom = 2
End Enum

'--- for WideCharToMultiByte
Private Const CP_UTF8 As Long = 65001

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private m_vLastError                As Variant
Private m_uPeekArray                As UcsSafeArraySingleDimension
Private m_aPeekBuffer()             As Integer

Private Type UcsSafeArraySingleDimension
    cDims       As Integer      '--- usually 1
    fFeatures   As Integer      '--- leave 0
    cbElements  As Long         '--- bytes per element (2-int, 4-long)
    cLocks      As Long         '--- leave 0
    pvData      As Long         '--- ptr to data
    cElements   As Long         '--- UBound + 1
    lLbound     As Long         '--- LBound
End Type

Public Function PushError(Optional vLocalErr As Variant) As Variant
    vLocalErr = Array(Err.Number, Err.Source, Err.Description, Erl)
    m_vLastError = vLocalErr
    PushError = vLocalErr
End Function

Public Function PopRaiseError(sFunction As String, sModule As String, Optional vLocalErr As Variant)
    If Not IsMissing(vLocalErr) Then
        m_vLastError = vLocalErr
    End If
    Err.Raise m_vLastError(0), sModule & "." & sFunction & vbCrLf & m_vLastError(1), m_vLastError(2)
End Function

Public Sub PopPrintError(sFunction As String, sModule As String, vLocalErr As Variant)
    If Not IsMissing(vLocalErr) Then
        m_vLastError = vLocalErr
    End If
    Debug.Print sModule & "." & sFunction & ": " & m_vLastError(2)
End Sub

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = ApiEmptyByteArray
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    FromUtf8Array = String$(2 * UBound(baText), 0)
    lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
    FromUtf8Array = Left$(FromUtf8Array, lSize)
End Function

Public Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Public Function RollingHash(ByVal lPtr As Long, ByVal lSize As Long) As Long
    Dim lIdx            As Long
    
    If lPtr = 0 Then
        Call CopyMemory(ByVal ArrPtr(m_aPeekBuffer), 0&, 4)
        m_uPeekArray.cDims = 0
        Exit Function
    ElseIf m_uPeekArray.cDims = 0 Then
        With m_uPeekArray
            .cDims = 1
            .cbElements = 2
        End With
        Call CopyMemory(ByVal ArrPtr(m_aPeekBuffer), VarPtr(m_uPeekArray), 4)
    End If
    m_uPeekArray.pvData = lPtr
    m_uPeekArray.cElements = lSize
    For lIdx = 0 To lSize - 1
        RollingHash = (RollingHash * 263 + m_aPeekBuffer(lIdx)) And &H3FFFFF
    Next
End Function

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    AssignVariant RetVal, pCol.Item(Index)
    SearchCollection = True
QH:
End Function

Public Function RemoveCollection(pCol As Collection, Index As Variant)
    On Error GoTo QH
    pCol.Remove Index
QH:
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Public Function GetModuleInstance(sModuleName As String, sInstanceName As String, Optional DebugID As String) As String
    If LenB(sInstanceName) <> 0 And LenB(DebugID) <> 0 Then
        GetModuleInstance = sModuleName & "(" & sInstanceName & ", " & DebugID & ")"
    ElseIf LenB(sInstanceName) <> 0 Or LenB(DebugID) <> 0 Then
        GetModuleInstance = sModuleName & "(" & Zn(sInstanceName, DebugID) & ")"
    Else
        GetModuleInstance = sModuleName
    End If
End Function

Public Function Zn(sText As String, Optional IfEmptyString As Variant = Null) As Variant
    Zn = IIf(LenB(sText) = 0, IfEmptyString, sText)
End Function

Public Function SystemIconFont() As StdFont
    Set SystemIconFont = New StdFont
End Function

#End If ' ImplUseShared


Private Sub pvClose(nFile As Integer)
    On Error GoTo EH
    If nFile <> 0 Then
        Close nFile
    End If
EH:
    nFile = 0
End Sub

Public Function ReadBinaryFile(sFile As String) As Byte()
    Const FUNC_NAME     As String = "ReadBinaryFile"
    Dim baBuffer()      As Byte
    Dim nFile           As Integer
    Dim vErr            As Variant
    
    On Error GoTo EH
    baBuffer = ApiEmptyByteArray()
    nFile = FreeFile
    Open sFile For Binary Access Read As nFile
    If LOF(nFile) > 0 Then
        ReDim baBuffer(0 To LOF(nFile) - 1) As Byte
        Get nFile, , baBuffer
    End If
    pvClose nFile
    ReadBinaryFile = baBuffer
    Exit Function
EH:
    PushError vErr
    pvClose nFile
    PopRaiseError FUNC_NAME & "(sFile=" & sFile & ")", MODULE_NAME, vErr
End Function

Public Sub WriteBinaryFile(sFile As String, baBuffer() As Byte)
    Const FUNC_NAME     As String = "WriteBinaryFile"
    Dim nFile           As Integer
    Dim vErr            As Variant
    
    On Error GoTo EH
    If InStrRev(sFile, "\") > 1 Then
        MkPath Left$(sFile, InStrRev(sFile, "\") - 1)
    End If
    DeleteFile sFile
    nFile = FreeFile
    Open sFile For Binary Access Write As nFile
    If Peek(ArrPtr(baBuffer)) <> 0 Then
        If UBound(baBuffer) >= 0 Then
            Put nFile, , baBuffer
        End If
    End If
    pvClose nFile
    Exit Sub
EH:
    PushError vErr
    pvClose nFile
    PopRaiseError FUNC_NAME & "(sFile=" & sFile & ")", MODULE_NAME, vErr
End Sub

Public Sub WriteTextFile(sFile As String, sText As String)
    Const FUNC_NAME     As String = "WriteTextFile"
    Dim nFile           As Integer
    Dim vErr            As Variant
    
    On Error GoTo EH
    If InStrRev(sFile, "\") > 1 Then
        MkPath Left$(sFile, InStrRev(sFile, "\") - 1)
    End If
    nFile = FreeFile
    Open sFile For Output As nFile
    Print #nFile, sText
    pvClose nFile
    Exit Sub
EH:
    PushError vErr
    pvClose nFile
    PopRaiseError FUNC_NAME & "(sFile=" & sFile & ")", MODULE_NAME, vErr
End Sub

Public Function FileAttr(sFile As String) As VbFileAttribute
    FileAttr = GetFileAttributes(sFile)
    If FileAttr = -1 Then
        FileAttr = &H8000
    End If
End Function

Public Function MkPath(sPath As String, Optional sError As String) As Boolean
    Const FUNC_NAME     As String = "MkPath"
    Dim vErr            As Variant
    
    On Error GoTo EH
    MkPath = (FileAttr(sPath) And vbDirectory) <> 0
    If Not MkPath Then
        If ApiCreateDirectory(sPath, 0) = 0 Then
            sError = GetSystemMessage(Err.LastDllError)
        End If
        MkPath = (FileAttr(sPath) And vbDirectory) <> 0
        If Not MkPath And InStrRev(sPath, "\") <> 0 Then
            MkPath Left$(sPath, InStrRev(sPath, "\") - 1)
            Call ApiCreateDirectory(sPath, 0)
            MkPath = (FileAttr(sPath) And vbDirectory) <> 0
        End If
    End If
    Exit Function
EH:
    PushError vErr
    PopRaiseError FUNC_NAME & "(sPath=" & sPath & ")", MODULE_NAME, vErr
End Function

Public Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim ret As Long
   
    GetSystemMessage = Space$(2000)
    ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If ret > 2 Then
        If Mid$(GetSystemMessage, ret - 1, 2) = vbCrLf Then
            ret = ret - 2
        End If
    End If
    GetSystemMessage = Left$(GetSystemMessage, ret)
End Function

Public Function DeleteFile(sFileName As String) As Boolean
    Call ApiDeleteFile(sFileName)
End Function

Public Function DesignDumpMemory(ByVal lPtr As Long, ByVal lSize As Long, Optional ByVal AddrPadding As Long = -1) As String
    Dim lIdx            As Long
    Dim sHex            As String
    Dim sChar           As String
    Dim lValue          As Long
    Dim cResult         As Collection
    Dim sPrefix         As String
    
    Set cResult = New Collection
    If lSize > 1 Then
        lIdx = Int(Log(lSize - 1) / Log(16) + 1)
    Else
        lIdx = 1
    End If
    If AddrPadding < 0 And AddrPadding < lIdx Then
        AddrPadding = lIdx
    End If
    For lIdx = 0 To ((lSize + 15) \ 16) * 16
        If lIdx < lSize Then
            If IsBadReadPtr(UnsignedAdd(lPtr, lIdx), 1) = 0 Then
                Call CopyMemory(lValue, ByVal UnsignedAdd(lPtr, lIdx), 1)
                sHex = sHex & Right$("00" & Hex$(lValue), 2) & " "
                If AddrPadding > 0 Then
                    If lValue >= 32 Then
                        sChar = sChar & Chr$(lValue)
                    Else
                        sChar = sChar & "."
                    End If
                End If
            Else
                sHex = sHex & "?? "
                If AddrPadding > 0 Then
                    sChar = sChar & "."
                End If
            End If
        Else
            sHex = sHex & "   "
        End If
        If ((lIdx + 1) Mod 16) = 0 Then
            If AddrPadding > 0 Then
                sPrefix = Right$(String(AddrPadding, "0") & Hex$(lIdx - 15), AddrPadding) & ": "
            End If
            cResult.Add RTrim$(sPrefix & sHex & " " & sChar)
            sHex = vbNullString
            sChar = vbNullString
        End If
    Next
    DesignDumpMemory = ConcatCollection(cResult, vbCrLf)
End Function

Public Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

'重复
Public Function UnsignedAdd2(ByVal Start As Long, ByVal Incr As Long) As Long
    UnsignedAdd2 = ((Start Xor &H80000000) + Incr) Xor &H80000000
End Function

Public Function ToArray(oCol As Collection) As Variant
    Const FUNC_NAME     As String = "ToArray"
    Dim vRetVal         As Variant
    Dim lIdx            As Long
    
    On Error GoTo EH
    If oCol.Count > 0 Then
        ReDim vRetVal(0 To oCol.Count - 1) As Variant
        For lIdx = 0 To UBound(vRetVal)
            vRetVal(lIdx) = oCol(lIdx + 1)
        Next
        ToArray = vRetVal
    Else
        ToArray = Array()
    End If
    Exit Function
EH:
    PopRaiseError FUNC_NAME, MODULE_NAME, PushError
End Function

Public Function CLngLng(vValue As Variant) As Variant
    Call VariantChangeType(CLngLng, vValue, 0, VT_I8)
End Function

Public Function ToLngLng(ByVal lLoDWord As Long, ByVal lHiDWord As Long) As Variant
    Call VariantChangeType(ToLngLng, ToLngLng, 0, VT_I8)
    Call CopyMemory(ByVal VarPtr(ToLngLng) + 8, lLoDWord, 4)
    Call CopyMemory(ByVal VarPtr(ToLngLng) + 12, lHiDWord, 4)
End Function

Public Function GetLoDWord(llValue As Variant) As Long
    Call CopyMemory(GetLoDWord, ByVal VarPtr(llValue) + 8, 4)
End Function

Public Function GetHiDWord(llValue As Variant) As Long
    Call CopyMemory(GetHiDWord, ByVal VarPtr(llValue) + 12, 4)
End Function

Public Function FormatXmlIndent(vDomOrString As Variant, sResult As String) As Boolean
    Dim oWriter As Object ' MSXML2.MXXMLWriter

    On Error GoTo QH
    Set oWriter = CreateObject("MSXML2.MXXMLWriter")
    oWriter.omitXMLDeclaration = True
    oWriter.Indent = True
    With CreateObject("MSXML2.SAXXMLReader")
        Set .contentHandler = oWriter
        '--- keep CDATA elements
        .putProperty "http://xml.org/sax/properties/lexical-handler", oWriter
        .Parse vDomOrString
    End With
    sResult = oWriter.Output
    '--- success
    FormatXmlIndent = True
    Exit Function
QH:
End Function



Public Function Save2ExcelB(sFile As String, mGrid As MSHFlexGrid, lCols As Integer, lRows As Long, Optional sInfo As String = "") As Boolean
    Dim oStyle() As cBiff12CellStyle
    Dim lIdx As Long
    Dim lRow As Long
    Dim dblTimer As Single
    Dim baBuffer() As Byte
    Dim i As Integer

    On Error GoTo EH
    dblTimer = Timer
    ReDim oStyle(0 To lCols) As cBiff12CellStyle
    '用于设置全局样式
    For lIdx = 0 To lCols - 1
        Set oStyle(lIdx) = New cBiff12CellStyle
        With oStyle(lIdx)
            .FontName = "宋体"
            .FontSize = 10 '+lIdx
            .Bold = False
            .BorderBottomColor = vbRed
            .BorderLeftColor = vbRed
            .BorderRightColor = vbRed
            .BorderTopColor = vbRed
            .VertAlign = ucsValMiddle
        End With
    Next
    'oStyle(0).WrapText = True
    oStyle(0).FontSize = 9
    'oStyle(0).HorAlign = ucsHalLeft
    oStyle(0).VertAlign = ucsValMiddle
    'oStyle(0).Format = "G/通用格式" ' "@" '设置A1单元格为文本格式
    
    'oStyle(5).WrapText = True
    oStyle(5).Bold = True
    oStyle(5).HorAlign = ucsHalCenter
    oStyle(5).VertAlign = ucsValMiddle
    'oStyle(5).BorderLeftColor = red
    'oStyle(5).Format = "0.00"
    
    Dim j As Integer
    With New cBiff12Writer
        '--- note: Excel's Biff12 clipboard reader cannot handle shared-strings table
        '使用 False 速度更快、文件更小
        .Init lCols, False, , "导出"
        '设置列宽必须在前面
        For i = 0 To lCols - 1
            .ColWidth(i) = IIf(mGrid.ColWidth(i) < 10, 1, mGrid.ColWidth(i) * 3)
        Next i
        Dim h As Integer
        h = 0
        For lRow = 0 To lRows - 1
            j = 0
            For lIdx = 0 To .ColCount - 1
                '跳过隐藏行，待处理
                'If mGrid.ColWidth(lIdx) > 10 Then
                    '.ColWidth(j) = mGrid.ColWidth(lIdx) * 3
                    'IIf(lRow = 18, 1000, 326)  '这里可以设置行高 mGrid.RowHeight(lRow) 可以设置每行高度，但会慢很多，实际上没必要
                    .AddRow lRow, mGrid.RowHeight(lRow)
                    .AddStringCell lIdx, mGrid.TextMatrix(lRow, lIdx), IIf(lRow = 0, oStyle(5), oStyle(0))
                    '这行必须在后面
                    j = j + 1
                    h = j
                'End If
            Next
            'If lRow = 1 Then
                '.MergeCells 7, 2, 3
                'If (FileAttr(App.Path & "\1.jpg") And vbArchive) <> 0 Then
                    '方法1
                    '.AddImage 7, ReadBinaryFile(App.Path & "\1.jpg"), 0, 0, 8839200, 8717450
                    '方法2
                    '.AddImage 7, SaveToPng(picAvatar.Picture), 0, 0, picAvatar.Picture.Width , picAvatar.Picture.Height
                'End If
            'End If
        Next
        
        If sInfo <> "" Then
            .AddRow  'lRow, 326 '这里可以设置行高
            .AddStringCell 0, sInfo, oStyle(0)
            For i = 1 To h - 1
                .AddStringCell i, "", oStyle(0)
            Next i
        End If
        .Flush
        .SaveToFile baBuffer
        WriteBinaryFile sFile, baBuffer
        
        'GetForegroundWindow()：在窗体中可以使用 me.hWnd，在模块中就要使用这个了。
        If OpenClipboard(GetForegroundWindow()) = 0 Then Err.Raise vbObjectError, , "打开剪切板错误"
        Call EmptyClipboard
        SetTextData "Biff12 Explorer"
        SetBinaryData AddFormat("Biff12"), baBuffer
        Call CloseClipboard
    End With
    Save2ExcelB = True
    'MsgBox "文件保存在： " & vbCrLf & vbCrLf & sFile & vbCrLf & vbCrLf & "耗时：" & Format$(Timer - dblTimer, "0.000") & "秒", 64
    Exit Function
EH:
    MsgBox Error & Err.Source, vbCritical, "Save2ExcelB"
End Function

Public Function SaveToPng(oPic As StdPicture) As Byte()
    With New cDibSection
        .LoadFromPicture oPic
        SaveToPng = .SaveToByteArray("image/png")
    End With
End Function

Private Function AddFormat(ByVal sName As String) As Long
    Dim wFormat As Long

    wFormat = RegisterClipboardFormat(sName & Chr$(0))
    If (wFormat > &HC000&) Then
        AddFormat = wFormat
    End If
End Function

Public Function SetBinaryData(ByVal lFormatId As Long, bData() As Byte) As Boolean
    Dim lSize           As Long
    Dim hMem            As Long
    Dim lPtr            As Long

    lSize = UBound(bData) - LBound(bData) + 1
    hMem = GlobalAlloc(GMEM_DDESHARE Or GMEM_MOVEABLE, lSize)
    If hMem = 0 Then
        GoTo QH
    End If
    lPtr = GlobalLock(hMem)
    If lPtr = 0 Then
        GoTo QH
    End If
    Call CopyMemory(ByVal lPtr, bData(LBound(bData)), lSize)
    Call GlobalUnlock(hMem)
    If SetClipboardData(lFormatId, hMem) = 0 Then
        GoTo QH
    End If
    SetBinaryData = True
QH:
End Function

Public Function SetTextData(sText As String) As Boolean
    Dim baData() As Byte
    
    If LenB(sText) > 0 Then
        baData = StrConv(sText & vbNullChar, vbFromUnicode)
        SetTextData = SetBinaryData(vbCFText, baData)
    End If
End Function
