Attribute VB_Name = "mdOpenSaveDialog"
Option Explicit
Private Const MODULE_NAME As String = "mdOpenSaveDialog"

Public Enum UcsOpenSaveDialogType
    ucsOsdOpen
    ucsOsdSave
End Enum

Public Enum UcsOpenSaveDirectoryType
    ucsOdtPersonal = &H5                         ' My Documents
    ucsOdtMyMusic = &HD                          ' "My Music" folder
    ucsOdtAppData = &H1A                         ' Application Data, new for NT4
    ucsOdtLocalAppData = &H1C                    ' non roaming, user\Local Settings\Application Data
    ucsOdtInternetCache = &H20
    ucsOdtCookies = &H21
    ucsOdtHistory = &H22
    ucsOdtCommonAppData = &H23                   ' All Users\Application Data
    ucsOdtWindows = &H24                         ' GetWindowsDirectory()
    ucsOdtSystem = &H25                          ' GetSystemDirectory()
    ucsOdtProgramFiles = &H26                    ' C:\Program Files
    ucsOdtMyPictures = &H27                      ' My Pictures, new for Win2K
    ucsOdtProgramFilesCommon = &H2B              ' C:\Program Files\Common
    ucsOdtCommonDocuments = &H2E                 ' All Users\Documents
    ucsOdtResources = &H38                       ' %windir%\Resources\, For theme and other windows resources.
    ucsOdtResourcesLocalized = &H39              ' %windir%\Resources\<LangID>, for theme and other windows specific resources.
    ucsOdtCommonAdminTools = &H2F                ' All Users\Start Menu\Programs\Administrative Tools
    ucsOdtAdminTools = &H30                      ' <user name>\Start Menu\Programs\Administrative Tools
    ucsOdtFlagCreate = &H8000&                   ' new for Win2K, or this in to force creation of folder
End Enum

Private Const OFN_OVERWRITEPROMPT       As Long = &H2&
Private Const OFN_HIDEREADONLY          As Long = &H4&
Private Const OFN_EXTENSIONDIFFERENT    As Long = &H400
Private Const OFN_CREATEPROMPT          As Long = &H2000&
Private Const OFN_EXPLORER              As Long = &H80000
Private Const OFN_LONGNAMES             As Long = &H200000
Private Const OFN_ENABLESIZING          As Long = &H800000
Private Const CDERR_DIALOGFAILURE       As Long = &HFFFF&
Private Const CDERR_STRUCTSIZE          As Long = &H1
Private Const CDERR_INITIALIZATION      As Long = &H2
Private Const CDERR_NOTEMPLATE          As Long = &H3
Private Const CDERR_NOHINSTANCE         As Long = &H4
Private Const CDERR_LOADSTRFAILURE      As Long = &H5
Private Const CDERR_FINDRESFAILURE      As Long = &H6
Private Const CDERR_LOADRESFAILURE      As Long = &H7
Private Const CDERR_LOCKRESFAILURE      As Long = &H8
Private Const CDERR_MEMALLOCFAILURE     As Long = &H9
Private Const CDERR_MEMLOCKFAILURE      As Long = &HA
Private Const CDERR_NOHOOK              As Long = &HB
Private Const FNERR_SUBCLASSFAILURE     As Long = &H3001
Private Const FNERR_INVALIDFILENAME     As Long = &H3002
Private Const FNERR_BUFFERTOOSMALL      As Long = &H3003

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (lpofn As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal l As Long)
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hWnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal szPath As String) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Type OPENFILENAME
    lStructSize         As Long     ' size of type/structure
    hWndOwner           As Long     ' Handle of owner window
    hInstance           As Long
    lpstrFilter         As Long     ' Filters used to select files
    lpstrCustomFilter   As Long
    nMaxCustomFilter    As Long
    nFilterIndex        As Long     ' index of Filter to start with
    lpstrFile           As Long     ' Holds filepath and name
    nMaxFile            As Long     ' Maximum Filepath and name length
    lpstrFileTitle      As Long     ' Filename
    nMaxFileTitle       As Long     ' Max Length of filename
    lpstrInitialDir     As Long     ' Starting Directory
    lpstrTitle          As Long     ' Title of window
    Flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As Long
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As Long
    pvReserved          As Long
    dwReserved          As Long
    FlagsEx             As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
    szCSDVersion        As String * 128
End Type

Private Const STR_ERROR                 As String = "打开文件错误(%1)"
Private Const STR_FNERR_INVALIDFILENAME As String = "无效的文件名(FNERR_INVALIDFILENAME)"
Private Const STR_FNERR_BUFFERTOOSMALL  As String = "缓冲区太小(FNERR_BUFFERTOOSMALL)"

Private Sub RaiseError(sFunction As String)
    PopRaiseError sFunction, MODULE_NAME, PushError
End Sub

Public Function ShowOpenSaveDialog(sFileName As String, sFilter As String, Optional sInitialDir As String, _
            Optional ByVal hWndOwner As Long, Optional sDefaultExt As String, Optional sTitle As String, _
            Optional ByVal eType As UcsOpenSaveDialogType = ucsOsdOpen) As Boolean
    Const FUNC_NAME     As String = "ShowOpenSaveDialog"
    Dim m_uOFN                  As OPENFILENAME
    Dim m_sFilter               As String
    Dim m_sDefExt               As String
    Dim m_sTitle                As String
    Dim m_sBufferCustomFilter   As String
    Dim m_sBufferFile           As String
    Dim m_sBufferInitialDir     As String
    Dim bRetry          As Boolean
    Dim sError          As String

    On Error GoTo EH
    m_sTitle = StrConv(sTitle, vbFromUnicode)
    m_sBufferCustomFilter = String(1024, 0)
    m_sBufferFile = String(1024, 0)
    m_sBufferInitialDir = String(1024, 0)
    With m_uOFN
        If OsVersion >= 500 Then
            .lStructSize = Len(m_uOFN)
        Else
            .lStructSize = Len(m_uOFN) - 12
        End If
        .hWndOwner = hWndOwner
        .lpstrCustomFilter = StrPtr(m_sBufferCustomFilter)
        .nMaxCustomFilter = Len(m_sBufferCustomFilter)
        .lpstrFile = StrPtr(m_sBufferFile)
        .nMaxFile = Len(m_sBufferFile)
        .lpstrTitle = StrPtr(m_sTitle)
        .lpstrInitialDir = StrPtr(m_sBufferInitialDir)
    End With
Retry:
    If eType = ucsOsdOpen And InStrRev(sFileName, "\") > 0 Then
        sInitialDir = Left(sFileName, InStrRev(sFileName, "\") - 1)
        sFileName = Mid(sFileName, InStrRev(sFileName, "\") + 1)
    End If
    If StrPtr(sFileName) <> 0 Then
        Call CopyMemory(ByVal m_uOFN.lpstrFile, ByVal sFileName, Len(sFileName) + 1)
    Else
        Call CopyMemory(ByVal m_uOFN.lpstrFile, ByVal "", Len(sFileName) + 1)
    End If
    m_sFilter = StrConv(Replace(sFilter, "|", vbNullChar), vbFromUnicode)
    If LenB(m_sFilter) <> 0 Then
        m_uOFN.lpstrFilter = StrPtr(m_sFilter)
    Else
        m_uOFN.lpstrFilter = 0
    End If
    If LenB(sInitialDir) = 0 Then
        sInitialDir = GetSpecialFolder(ucsOdtPersonal)
    End If
    If StrPtr(sInitialDir) <> 0 Then
        Call CopyMemory(ByVal m_uOFN.lpstrInitialDir, ByVal sInitialDir, Len(sInitialDir) + 1)
    Else
        Call CopyMemory(ByVal m_uOFN.lpstrInitialDir, ByVal "", Len(sInitialDir) + 1)
    End If
    If LenB(sFilter) <> 0 Then
        m_uOFN.nFilterIndex = 1
    End If
    If LenB(sDefaultExt) <> 0 Then
        m_sDefExt = StrConv(sDefaultExt, vbFromUnicode)
        If LenB(m_sDefExt) <> 0 Then
            m_uOFN.lpstrDefExt = StrPtr(m_sDefExt)
        Else
            m_uOFN.lpstrDefExt = 0
        End If
    End If
    If eType = ucsOsdOpen Then
        m_uOFN.Flags = OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_HIDEREADONLY Or OFN_EXTENSIONDIFFERENT Or OFN_EXPLORER Or OFN_ENABLESIZING '--- Or OFN_ENABLEHOOK
        If GetOpenFileName(m_uOFN) Then
            ShowOpenSaveDialog = True
        End If
    Else
        m_uOFN.Flags = OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY Or OFN_EXTENSIONDIFFERENT Or OFN_EXPLORER Or OFN_ENABLESIZING ' Or OFN_ENABLEHOOK
        If GetSaveFileName(m_uOFN) Then
            ShowOpenSaveDialog = True
        End If
    End If
    If Not ShowOpenSaveDialog Then
        If Not bRetry And CommDlgExtendedError() = FNERR_INVALIDFILENAME Then
            bRetry = True
            sFileName = vbNullString
            GoTo Retry
        End If
        sError = pvTranslateError(CommDlgExtendedError())
    End If
    If LenB(sError) Then
        On Error GoTo 0
        Err.Raise vbObjectError, , sError
    End If
    sFileName = pvToString(m_uOFN.lpstrFile)
    Exit Function
EH:
    RaiseError FUNC_NAME
End Function

Public Function GetSpecialFolder(ByVal eType As UcsOpenSaveDirectoryType) As String
    GetSpecialFolder = String(1000, 0)
    Call SHGetFolderPath(0, eType Or ucsOdtFlagCreate, 0, 0, GetSpecialFolder)
    GetSpecialFolder = Left(GetSpecialFolder, InStr(GetSpecialFolder, Chr$(0)) - 1)
End Function

Public Property Get OsVersion() As Long
    Static lVersion     As Long
    Dim uVer            As OSVERSIONINFO
    
    If lVersion = 0 Then
        uVer.dwOSVersionInfoSize = Len(uVer)
        If GetVersionEx(uVer) Then
            lVersion = uVer.dwMajorVersion * 100 + uVer.dwMinorVersion
        End If
    End If
    OsVersion = lVersion
End Property

Private Function pvTranslateError(ByVal lRetVal As Long) As String
    Select Case lRetVal
    Case 0
    Case FNERR_INVALIDFILENAME
        pvTranslateError = STR_FNERR_INVALIDFILENAME
    Case FNERR_BUFFERTOOSMALL
        pvTranslateError = STR_FNERR_BUFFERTOOSMALL
    Case Else
        Select Case lRetVal
        Case CDERR_DIALOGFAILURE
            pvTranslateError = "CDERR_DIALOGFAILURE"
        Case CDERR_STRUCTSIZE
            pvTranslateError = "CDERR_STRUCTSIZE"
        Case CDERR_INITIALIZATION
            pvTranslateError = "CDERR_INITIALIZATION"
        Case CDERR_NOTEMPLATE
            pvTranslateError = "CDERR_NOTEMPLATE"
        Case CDERR_NOHINSTANCE
            pvTranslateError = "CDERR_NOHINSTANCE"
        Case CDERR_LOADSTRFAILURE
            pvTranslateError = "CDERR_LOADSTRFAILURE"
        Case CDERR_FINDRESFAILURE
            pvTranslateError = "CDERR_FINDRESFAILURE"
        Case CDERR_LOADRESFAILURE
            pvTranslateError = "CDERR_LOADRESFAILURE"
        Case CDERR_LOCKRESFAILURE
            pvTranslateError = "CDERR_LOCKRESFAILURE"
        Case CDERR_MEMALLOCFAILURE
            pvTranslateError = "CDERR_MEMALLOCFAILURE"
        Case CDERR_MEMLOCKFAILURE
            pvTranslateError = "CDERR_MEMLOCKFAILURE"
        Case CDERR_NOHOOK
            pvTranslateError = "CDERR_NOHOOK"
        Case FNERR_SUBCLASSFAILURE
            pvTranslateError = "FNERR_SUBCLASSFAILURE"
        Case Else
            pvTranslateError = lRetVal
        End Select
        pvTranslateError = Replace(STR_ERROR, "%1", pvTranslateError)
    End Select
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String(lstrlen(lPtr), Chr(0))
        Call CopyMemory(ByVal pvToString, ByVal lPtr, lstrlen(lPtr))
    End If
End Function

Public Sub FlexGridToCsv(mGrid As MSHFlexGrid, Filename As String, IsView As Boolean, Optional Delimiter As String = ",", Optional sOther As String = "")
    On Error GoTo Err_Out
    Dim fh As Integer
    Dim FileIsOpen As Boolean
    Dim i As Integer
    Dim Buf As String, TempStr As String
    FileIsOpen = False
    
    fh = FreeFile()
    Open Filename For Output As fh
    FileIsOpen = True
    Buf = ""
    For i = 0 To mGrid.Cols - 1
        If Buf = "" Then
            Buf = """" & mGrid.TextMatrix(0, i) & """"
        Else
            If IsView = True Then
                If mGrid.ColWidth(i) > 10 Then
                    Buf = Buf & Delimiter & """" & mGrid.TextMatrix(0, i) & """"
                End If
            Else
                Buf = Buf & Delimiter & """" & mGrid.TextMatrix(0, i) & """"
            End If
        End If
    Next i
    Print #fh, Buf
    
    mGrid.RowSel = 1
    Dim ii As Long
    For ii = 0 To mGrid.Rows - 2
        Buf = ""
        For i = 0 To mGrid.Cols - 1
            If IsNull(mGrid.TextMatrix(mGrid.RowSel, i)) Then
                TempStr = ""
            Else
                TempStr = mGrid.TextMatrix(mGrid.RowSel, i)
            End If
            
            If Buf = "" Then
                '第一列是订单号、数字超过8位、直接加'号
                Buf = """'" & TempStr & """"
            Else
                If IsView = True Then
                    If mGrid.ColWidth(i) > 10 Then
                        Buf = Buf & Delimiter & IIf(MyIsDate(TempStr), """'", """") & TempStr & """"
                    End If
                Else
                    Buf = Buf & Delimiter & IIf(MyIsDate(TempStr), """'", """") & TempStr & """"
                End If
            End If
        Next i
        If mGrid.RowSel <= mGrid.Rows - 2 Then
            mGrid.RowSel = mGrid.RowSel + 1
        End If
        Print #fh, Buf
    Next ii
    Print #fh, sOther
    Close fh
    Exit Sub
Err_Out:
    If FileIsOpen Then Close fh
    MsgBox "导文件错误： " & Error, vbOKOnly, "错误"
End Sub

Public Function MyIsDate(DateStr As String) As Boolean
    Dim S As String
    If Len(DateStr) < 6 Then Exit Function
    S = Left(DateStr, 4) & "-" & Mid(DateStr, 5, 2) '& "-" & Mid(DateStr, 7, 2)
    If (IsDate(S) = True And Len(DateStr) = 7) Or (IsDate(DateStr) = True And Len(DateStr) = 7) Or (IsNumeric(DateStr) = True And Len(DateStr) > 8) Then
        MyIsDate = True
    Else
        MyIsDate = False
    End If
End Function

