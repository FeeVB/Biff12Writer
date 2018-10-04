VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}#1.0#0"; "Inked.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Biff12 Explorer"
   ClientHeight    =   8235
   ClientLeft      =   1980
   ClientTop       =   2325
   ClientWidth     =   15405
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   15405
   Begin VB.PictureBox picAvatar 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   630
      Left            =   375
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   630
      ScaleWidth      =   795
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6375
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   435
      TabIndex        =   0
      Top             =   2745
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8250
      Left            =   3390
      TabIndex        =   3
      Top             =   0
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   14552
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5808
      Left            =   84
      TabIndex        =   1
      Top             =   84
      Width           =   5304
      _ExtentX        =   9340
      _ExtentY        =   10239
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin INKEDLibCtl.InkEdit RichTextBox1 
      Height          =   5724
      Left            =   5964
      OleObjectBlob   =   "Form1.frx":23D2
      TabIndex        =   2
      Top             =   84
      Width           =   4800
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1995
      Top             =   6795
      _ExtentX        =   979
      _ExtentY        =   979
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   14549247
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":261B
            Key             =   "doc"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":2679
            Key             =   "folder"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMain 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit"
         Index           =   3
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Debug"
      Index           =   1
      Begin VB.Menu mnuDebug 
         Caption         =   "Minimal save"
         Index           =   0
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Test writer"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Biff12Writer (c) 2017 by wqweto@gmail.com
' A VB6 library for consuming/producing BIFF12 (.xlsb) spreadsheets
Option Explicit
DefObj A-Z


Private Const STR_DUMMY             As String = "$dummy"
Private Const STR_OPEN_FILTER As String = "Excel 文件(*.xlsb;*.xlsx)|*.xlsb;*.xlsx|All files (*.*)|*.*"
Private Const STR_OPEN_TITLE        As String = "Load BIFF12 file"

Private m_oZip As cZipArchive
Private m_lSeqNo As Long

Private Enum UcsMenuEnum
    ucsMnuNew = 0
    ucsMnuOpen = 1
    ucsMnuExit = 3
    ucsMnuMinSave = 0
    ucsMnuTestWriter = 1
End Enum

Private Sub pvLoadBiff12File(oTree As TreeView, sFile As String)
    Dim lIdx            As Long
    Dim sName           As String
    Dim oNode           As ComctlLib.Node

    Set m_oZip = New cZipArchive
    If Not m_oZip.OpenArchive(sFile) Then
        MsgBox "Error opening archive. " & m_oZip.LastError, vbExclamation
        GoTo QH
    End If
    oTree.Nodes.Clear
    oTree.Nodes.Add(, , "Root", Mid$(sFile, InStrRev(sFile, "\") + 1)).Expanded = True
    For lIdx = 0 To m_oZip.FileCount - 1
        sFile = m_oZip.FileInfo(lIdx)(0)
        sName = Mid$(sFile, InStrRev(sFile, "\") + 1)
        If LenB(sName) <> 0 Then
            Set oNode = oTree.Nodes.Add(pvGetParentKey("Root\" & sFile), tvwChild, sFile, Mid$(sFile, InStrRev(sFile, "\") + 1), "doc")
            If LCase$(Right$(sFile, 4)) = ".bin" And InStr(sFile, "printerSettings") = 0 Then
                oTree.Nodes.Add sFile, tvwChild, sFile & STR_DUMMY
            Else
                oNode.Tag = STR_DUMMY
                '--- immediate load
                pvDelayLoad m_oZip, oTree, oNode
            End If
        End If
    Next
    Set oTree.SelectedItem = oTree.Nodes(1)
    TreeView1_NodeClick oTree.SelectedItem
    Caption = oTree.Nodes(1).Text & " - Biff12 Explorer"
QH:
End Sub

Private Function pvLoadBinFile(oBin As cBiff12Part, oTree As TreeView, sRoot As String) As Boolean
    Dim eRecID          As UcsBiff12RecortTypeEnum
    Dim lRecSize        As Long
    Dim lPos            As Long
    Dim cStack          As Collection
    Dim sKey            As String
    Dim sName           As String
    Dim oNode           As ComctlLib.Node
    Dim sPrevSel        As String
    Dim dblTimer        As Double

    On Error GoTo EH
    If Not oTree.SelectedItem Is Nothing Then
        sPrevSel = oTree.SelectedItem.Key
    End If
    dblTimer = Timer
    Set cStack = New Collection
    cStack.Add sRoot
    eRecID = oBin.ReadVarDWord()
    lRecSize = oBin.ReadVarDWord()
    Do While eRecID <> -1
        m_lSeqNo = m_lSeqNo + 1
        sKey = "#" & m_lSeqNo
        sName = GetBrtName(eRecID) & ", pos=" & lPos & IIf(lRecSize <> 0, ", size=" & lRecSize, vbNullString)
        lPos = oBin.Position + lRecSize
        Set oNode = oTree.Nodes.Add(cStack(cStack.Count), tvwChild, sKey, sName)
        If InStr(1, sName, "Begin", vbBinaryCompare) Then
            cStack.Add sKey
            oNode.Expanded = True
        ElseIf InStr(1, sName, "End", vbBinaryCompare) Then
            cStack.Remove cStack.Count
        End If
        oNode.Tag = GetBrtData(eRecID, lRecSize, oBin)
        If Not IsArray(oNode.Tag) And lRecSize > 0 Then
            oNode.Text = oNode.Text & " [raw]"
        End If
        oBin.Position = lPos
        '--- unknown record, possibly structured storage file (like printerSettings1.bin)
        If Left$(sName, 2) = "0x" Then
            Exit Do
        End If
        If dblTimer + 5 < Timer Then
            If MsgBox("Too many nodes!" & vbCrLf & vbCrLf & "Do you want to continue?", vbQuestion Or vbYesNo) = vbYes Then
                dblTimer = 2 ^ 30
            Else
                Exit Do
            End If
        End If
        eRecID = oBin.ReadVarDWord()
        lRecSize = oBin.ReadVarDWord()
    Loop
    Set oTree.SelectedItem = oTree.Nodes(sRoot)
    If LenB(sPrevSel) <> 0 Then
        On Error Resume Next
        Set oTree.SelectedItem = oTree.Nodes(sPrevSel)
        On Error GoTo 0
    End If
    oTree.SelectedItem.EnsureVisible
    '--- success
    pvLoadBinFile = True
    Exit Function
EH:
    MsgBox Error, vbCritical
    Resume
End Function

Private Function pvGetParentKey(sFile As String) As String
    Dim lPos            As Long
    Dim lPrevPos        As Long

    lPos = InStr(1, sFile, "\")
    Do While lPos > 0
        If Not SearchCollection(TreeView1.Nodes, Left$(sFile, lPos - 1)) Then
            With TreeView1.Nodes.Add(Left$(sFile, lPrevPos - 1), tvwChild, Left$(sFile, lPos - 1), Mid$(sFile, lPrevPos + 1, lPos - lPrevPos - 1), "folder")
                .Expanded = True
            End With
        End If
        lPrevPos = lPos
        lPos = InStr(lPos + 1, sFile, "\")
    Loop
    pvGetParentKey = Left$(sFile, lPrevPos - 1)
End Function

Private Function pvDelayLoad(oZip As cZipArchive, oTree As TreeView, oNode As ComctlLib.Node) As Boolean
    Dim oStream         As cBiff12Part
    Dim oBin            As cBiff12Part
    Dim baContents()    As Byte
    Dim sXml            As String
    Dim bIsBinPart      As Boolean

    On Error GoTo EH
    If oNode.Image <> "doc" Then
        Exit Function
    End If
    If oNode.Children = 1 Then
        If oNode.Child.Key <> oNode.Key & STR_DUMMY Then
            Exit Function
        End If
        TreeView1.Nodes.Remove oNode.Key & STR_DUMMY
        bIsBinPart = True
    ElseIf oNode.Tag <> STR_DUMMY Then
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    Set oStream = New cBiff12Part
    If Not oZip.Extract(oStream, oNode.Key) Then
        MsgBox "Error extracting. " & oZip.LastError, vbExclamation
        GoTo QH
    End If
    If bIsBinPart Then
        Set oBin = New cBiff12Part
        oBin.Contents = oStream.Contents
        pvDelayLoad = pvLoadBinFile(oBin, oTree, oNode.Key)
    Else
        baContents = oStream.Contents
        If UBound(baContents) >= 0 Then
            If FormatXmlIndent(FromUtf8Array(baContents), sXml) Then
                oNode.Tag = sXml
            Else
                oNode.Tag = DesignDumpMemory(VarPtr(baContents(0)), UBound(baContents) + 1)
            End If
        End If
    End If
QH:
    Screen.MousePointer = vbDefault
    Exit Function
EH:
    MsgBox Error, vbCritical
    Resume
End Function

Private Function pvEnumTags(oNode As ComctlLib.Node, Optional ByVal lIndent As Long = -4, Optional cOutput As Collection) As Collection
    Dim vElem           As Variant
    Dim oChild          As ComctlLib.Node

    If cOutput Is Nothing Then
        Set cOutput = New Collection
    End If
    If lIndent >= 0 Then
        cOutput.Add Space$(lIndent) & "[" & IIf(oNode.Child Is Nothing, "-", IIf(oNode.Expanded, "-", "+")) & "] " & oNode.Text
    Else
        cOutput.Add oNode.Text
    End If
    lIndent = lIndent + 4
    If IsArray(oNode.Tag) Then
        For Each vElem In oNode.Tag
            cOutput.Add Space$(lIndent) & Replace(vElem, vbCrLf, vbCrLf & Space(lIndent))
        Next
    ElseIf LenB(oNode.Tag) <> 0 And oNode.Tag <> STR_DUMMY Then
        cOutput.Add Space$(lIndent) & Replace(oNode.Tag, vbCrLf, vbCrLf & Space(lIndent))
    End If
    If oNode.Expanded Then
        Set oChild = oNode.Child
        Do While Not oChild Is Nothing
            pvEnumTags oChild, lIndent, cOutput
            Set oChild = oChild.Next
        Loop
    End If
    Set pvEnumTags = cOutput
End Function

Private Sub pvTestMinimalSave(sFile As String)
    Dim uFont           As UcsBiff12BrtFontType
    Dim uFill           As UcsBiff12BrtFillType
    Dim uBorder         As UcsBiff12BrtBorderType
    Dim uXf             As UcsBiff12BrtXfType
    Dim oFile           As cBiff12Container
    Dim uBundle         As UcsBiff12BrtBundleShType
    Dim uWsProp         As UcsBiff12BrtWsPropType
    Dim uWsDim          As UcsBiff12UncheckedRfXType
    Dim uColInfo        As UcsBiff12BrtColInfoType
    Dim uRowHdr         As UcsBiff12BrtRowHdrType
    Dim oStrings        As cBiff12Part
'    Dim lPos            As Long
'    Dim lSize           As Long

    Set oFile = New cBiff12Container
    oFile.GetRelID oFile.WorkbookPart, oFile.SheetPart(1)

    ' STYLESHEET = BrtBeginStyleSheet [FMTS] [FONTS] [FILLS] [BORDERS] CELLSTYLEXFS CELLXFS STYLES DXFS TABLESTYLES [COLORPALETTE] FRTSTYLESHEET BrtEndStyleSheet
    With oFile.StylesPart
        .Output ucsBrtBeginStyleSheet

            .OutputCount ucsBrtBeginFonts, 1
                uFont.m_dyHeight = 220
                uFont.m_bls = 400
                uFont.m_bFamily = 2
                uFont.m_bCharSet = 204
                uFont.m_brtColor.m_xColorType = 3 * 2 + 1
                uFont.m_brtColor.m_index = 1
                uFont.m_brtColor.m_bAlpha = 255
                uFont.m_bFontScheme = 2
                uFont.m_name = "Calibri"
                .OutputBrtFont uFont
'                .WriteRecord ucsBrtACBegin, 6
'                .WriteDWord &HE020001
'                .WriteWord &H8000
'                    .Output ucsBrtKnownFonts
'                .Output ucsBrtACEnd
            .Output ucsBrtEndFonts

            .OutputCount ucsBrtBeginFills, 2
                uFill.m_fls = 0
                With uFill.m_brtColorFore
                    .m_xColorType = 1 * 2 + 1
                    .m_index = 64
                    .m_bAlpha = 255
                End With
                With uFill.m_brtColorBack
                    .m_xColorType = 1 * 2 + 1
                    .m_index = 65
                    .m_bRed = 255
                    .m_bGreen = 255
                    .m_bBlue = 255
                    .m_bAlpha = 255
                End With
                .OutputBrtFill uFill
                uFill.m_fls = &H11
                With uFill.m_brtColorFore
                    .m_xColorType = 1 * 2 + 1
                    .m_index = 64
                    .m_bAlpha = 255
                End With
                With uFill.m_brtColorBack
                    .m_xColorType = 1 * 2 + 1
                    .m_index = 65
                    .m_bRed = 255
                    .m_bGreen = 255
                    .m_bBlue = 255
                    .m_bAlpha = 255
                End With
                .OutputBrtFill uFill
            .Output ucsBrtEndFills

            .OutputCount ucsBrtBeginBorders, 1
                uBorder.m_blxfTop.m_brtColor.m_xColorType = 1
                uBorder.m_blxfBottom.m_brtColor.m_xColorType = 1
                uBorder.m_blxfLeft.m_brtColor.m_xColorType = 1
                uBorder.m_blxfRight.m_brtColor.m_xColorType = 1
                uBorder.m_blxfDiag.m_brtColor.m_xColorType = 1
                .OutputBrtBorder uBorder
            .Output ucsBrtEndBorders

'            .OutputCount ucsBrtBeginCellStyleXFs, 1
'                uXf.m_ixfeParent = -1
'                uXf.m_flags = &H1010
'                .OutputBrtXf uXf
'            .Output ucsBrtEndCellStyleXFs

            .OutputCount ucsBrtBeginCellXFs, 1
                uXf.m_ixfeParent = 0
                uXf.m_flags = &H1010
                .OutputBrtXf uXf
            .Output ucsBrtEndCellXFs

'            .OutputCount ucsBrtBeginStyles, 1
'                uStyle.m_grbitObj1 = 1
'                uStyle.m_iLevel = 255
'                uStyle.m_stName = "Normal"
'                .OutputBrtStyle uStyle
'            .Output ucsBrtEndStyles
'
'            .OutputCount ucsBrtBeginDXFs, 0
'            .Output ucsBrtEndDXFs

'            Const STR_TS_DEFLIST As String = "TableStyleMedium2"
'            Const STR_TS_DEFPIVOT As String = "PivotStyleLight16"
'            .WriteRecord ucsBrtBeginTableStyles, 4 + 4 + LenB(STR_TS_DEFLIST) + 4 + LenB(STR_TS_DEFPIVOT)
'            .WriteDWord 0
'            .WriteString STR_TS_DEFLIST
'            .WriteString STR_TS_DEFPIVOT
'            .Output ucsBrtEndTableStyles

        .Output ucsBrtEndStyleSheet
    End With

    ' WORKBOOK = BrtBeginBook [BrtFileVersion] [[BrtFileSharingIso] BrtFileSharing] [BrtWbProp] [ACABSPATH] [ACREVISIONPTR] [[BrtBookProtectionIso] BrtBookProtection] [BOOKVIEWS] BUNDLESHS [FNGROUP] [EXTERNALS] *BrtName [BrtCalcProp] [BrtOleSize] *(BrtUserBookView *FRT) [PIVOTCACHEIDS] [BrtWbFactoid] [SMARTTAGTYPES] [BrtWebOpt] *BrtFileRecover [WEBPUBITEMS] [CRERRS] FRTWORKBOOK BrtEndBook
    With oFile.WorkbookPart
        .Output ucsBrtBeginBook

'            lSize = 50
'            lPos = .WriteRecord(ucsBrtFileVersion, lSize)
'            .WriteGuid vbNullString
'            .WriteString "vb"
'            .WriteString "6"
'            .WriteString "6"
'            .WriteString "14420"
'            Debug.Assert lPos + lSize = .Position

'            uWbProp.m_flags = &H10020
'            uWbProp.m_dwThemeVersion = 153222
'            .OutputBrtWbProp uWbProp

'            .Output ucsBrtBeginBookViews
'                uBookView.m_dxWn = 30720
'                uBookView.m_dyWn = 13704
'                uBookView.m_iTabRatio = 600
'                uBookView.m_flags = &H78
'                .OutputBrtBookView uBookView
'            .Output ucsBrtEndBookViews

            .Output ucsBrtBeginBundleShs
                uBundle.m_hsState = 0
                uBundle.m_iTabID = 1
                uBundle.m_strRelID = oFile.GetRelID(oFile.WorkbookPart, oFile.SheetPart(1))
                uBundle.m_strName = "Sheet1"
                .OutputBrtBundleSh uBundle
            .Output ucsBrtEndBundleShs

        .Output ucsBrtEndBook
    End With

    ' SHAREDSTRINGS = BrtBeginSst *BrtSSTItem *FRT BrtEndSst
    Set oStrings = oFile.StringsPart
    oStrings.OutputCount2 ucsBrtBeginSst, 0, 0

    ' WORKSHEET = BrtBeginSheet [BrtWsProp] [BrtWsDim] [WSVIEWS2] [WSFMTINFO] *COLINFOS CELLTABLE [BrtSheetCalcProp] [[BrtSheetProtectionIso] BrtSheetProtection] *([BrtRangeProtectionIso] BrtRangeProtection) [SCENMAN] [AUTOFILTER] [SORTSTATE] [DCON] [USERSHVIEWS] [MERGECELLS] [BrtPhoneticInfo] *CONDITIONALFORMATTING [DVALS] *([ACUID] BrtHLink) [BrtPrintOptions] [BrtMargins] [BrtPageSetup] [HEADERFOOTER] [RWBRK] [COLBRK] *BrtBigName [CELLWATCHES] [IGNOREECS] [SMARTTAGS] [BrtDrawing] [BrtLegacyDrawing] [BrtLegacyDrawingHF] [BrtBkHim] [OLEOBJECTS] [ACTIVEXCONTROLS] [WEBPUBITEMS] [LISTPARTS] FRTWORKSHEET [ACUID] BrtEndSheet
    With oFile.SheetPart
        .Output ucsBrtBeginSheet

            uWsProp.m_flags = &H204C9
            uWsProp.m_brtcolorTab.m_index = 64
            uWsProp.m_rwSync = -1
            uWsProp.m_colSync = -1
            .OutputBrtWsProp uWsProp

            uWsDim.m_colLast = 2
            .OutputBrtWsDim uWsDim

            ' COLINFOS = BrtBeginColInfos 1*BrtColInfo BrtEndColInfos
            .Output ucsBrtBeginColInfos
                uColInfo.m_colLast = 2
                uColInfo.m_colDx = 1440
                .OutputBrtColInfo uColInfo
            .Output ucsBrtEndColInfos

             'CELLTABLE = BrtBeginSheetData *1048576([ACCELLTABLE] BrtRowHdr *16384CELL *FRT) BrtEndSheetData
            .Output ucsBrtBeginSheetData
                uRowHdr.m_rw = 0
                uRowHdr.m_miyRw = 288
                uRowHdr.m_ccolspan = 1
                ReDim uRowHdr.m_rgBrtColspan(0 To 0) As UcsBiff12BrtColSpanType
                uRowHdr.m_rgBrtColspan(0).m_colLast = 2
                .OutputBrtRowHdr uRowHdr

                .OutputCellIsst 0, 0, oStrings.SstGetIndex("Test")
'                .OutputCellBlank 1, 0
                .OutputCellIsst 2, 0, oStrings.SstGetIndex("aaa")
            .Output ucsBrtEndSheetData

'            .OutputCount ucsBrtBeginMergeCells, 0 ' MERGECELLS
'            .Output ucsBrtEndMergeCells

        .Output ucsBrtEndSheet
    End With

    oStrings.Output ucsBrtEndSst

'    oFile.AppPropsPart.XmlDocument.Load TEMP_FOLDER & "\Book3.xlsb\docProps\app.xml"
'    oFile.ThemePart.XmlDocument.Load TEMP_FOLDER & "\Book3.xlsb\xl\theme\theme1.xml"

    oFile.SaveToFile sFile
End Sub

Private Function pvTestBiff12Writer(sFile As String) As Boolean
    Const CLR_GREY      As Long = &HC0C0C0
    Dim oStyle()        As cBiff12CellStyle
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim dblTimer        As Single
    Dim baBuffer()      As Byte

    On Error GoTo EH
    dblTimer = Timer
    ReDim oStyle(0 To 5) As cBiff12CellStyle
    For lIdx = 0 To 5
        Set oStyle(lIdx) = New cBiff12CellStyle
        With oStyle(lIdx)
            .FontName = "宋体"
            .FontSize = 9 + lIdx
            .Bold = True
            .BorderLeftColor = CLR_GREY
            .BorderRightColor = CLR_GREY
        End With
    Next
    oStyle(0).BorderLeftColor = vbBlack
    oStyle(3).Format = "0,#.00"
    oStyle(4).BorderRightColor = vbBlack
    oStyle(4).ForeColor = vbRed
    oStyle(4).BackColor = CLR_GREY
    oStyle(4).WrapText = True
    With New cBiff12Writer
        '--- note: Excel's Biff12 clipboard reader cannot handle shared-strings table
        .Init 5, True, , "中国测试"
        For lRow = 0 To 10
            If lRow = 0 Then
                .MergeCells 0, 2, 3
                If (FileAttr(App.Path & "\1.jpg") And vbArchive) <> 0 Then
                    .AddImage 5, ReadBinaryFile(App.Path & "\1.jpg"), 0, 0, 883920, 871745
                End If
            End If
            For lIdx = 0 To .ColCount - 1
                With oStyle(lIdx)
                    .BorderTopColor = IIf(lRow = 0, vbBlack, CLR_GREY)
                    .BorderBottomColor = IIf(lRow = 2, vbBlack, CLR_GREY)
                End With
            Next
            .AddRow lRow
            .AddStringCell 0, "Test", oStyle(0)
            .AddStringCell 1, vbNullString, oStyle(1)
            .AddStringCell 2, "有地有要上", oStyle(2)
            .AddNumberCell 3, Round(lRow + Timer - 60000, 3), oStyle(3)
            .AddStringCell 4, lRow & " - " & Now, oStyle(4)
        Next
        .Flush
        .SaveToFile baBuffer
        WriteBinaryFile sFile, baBuffer
        
        'If OpenClipboard(Me.hWnd) = 0 Then Err.Raise vbObjectError, , "Cannot open clipboard"
        'Call EmptyClipboard
        SetTextData "Biff12 Explorer"
        'SetBinaryData AddFormat("Biff12"), baBuffer
        'Call CloseClipboard
    End With
    'MsgBox "Save of " & sFile & " complete in " & Format$(Timer - dblTimer, "0.000"), vbExclamation
    pvTestBiff12Writer = True
    Exit Function
EH:
    MsgBox Error, vbCritical
End Function

Private Sub Command1_Click()
    Save2ExcelB App.Path & "\test.xlsb", MSHFlexGrid1, MSHFlexGrid1.Cols, MSHFlexGrid1.Rows, "sadfsf " & Now
    'FlexGridToCsv Me.MSHFlexGrid1, App.Path & "\test.csv", False
    Unload Me
End Sub

Private Sub Form_Load()
    TreeView1.Move 84, TreeView1.Top, TreeView1.Width, ScaleHeight - TreeView1.Top - 84
    RichTextBox1.Move TreeView1.Left + TreeView1.Width + 84, TreeView1.Top, ScaleWidth - RichTextBox1.Left - 84, TreeView1.Height
    RichTextBox1.Width = ScaleWidth - RichTextBox1.Left - 84
    
    Dim i As Long, j As Long
    Me.MSHFlexGrid1.Rows = 3
    Me.MSHFlexGrid1.Cols = 7
    
    For i = 1 To 2
        For j = 0 To 6
            MSHFlexGrid1.TextMatrix(0, j) = "列" & j
            MSHFlexGrid1.TextMatrix(i, j) = i & j
        Next j
    Next i
    MSHFlexGrid1.ColWidth(0) = 2800
    MSHFlexGrid1.ColWidth(1) = 3600
    MSHFlexGrid1.ColWidth(2) = 0
    MSHFlexGrid1.ColWidth(3) = 1600
    MSHFlexGrid1.ColWidth(4) = 0
    MSHFlexGrid1.ColWidth(5) = 1800
    MSHFlexGrid1.ColWidth(6) = 3000
    MSHFlexGrid1.TextMatrix(1, 1) = "1234567890"
    MSHFlexGrid1.TextMatrix(1, 2) = "12345678901234567890"
    MSHFlexGrid1.TextMatrix(1, 6) = "12345.678"
    MSHFlexGrid1.TextMatrix(2, 0) = "2018-08-01"
    MSHFlexGrid1.TextMatrix(2, 1) = "2018-08"
    MSHFlexGrid1.TextMatrix(2, 2) = "2018-08-01 16:18:28"
End Sub

Private Sub mnuDebug_Click(Index As Integer)
    Select Case Index
    Case ucsMnuMinSave
        pvTestMinimalSave App.Path & "\output.xlsb"
    Case ucsMnuTestWriter
        If pvTestBiff12Writer(App.Path & "\output.xlsb") Then
            pvLoadBiff12File TreeView1, App.Path & "\output.xlsb"
        End If
    End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Dim sFile           As String
    Dim sInitDir        As String

    Select Case Index
    Case ucsMnuOpen
        If ShowOpenSaveDialog(sFile, STR_OPEN_FILTER, , Me.hWnd, , ucsOsdOpen) Then
            Screen.MousePointer = vbHourglass
            pvLoadBiff12File TreeView1, sFile
            Screen.MousePointer = vbDefault
        End If
    Case ucsMnuExit
        Unload Me
    End Select
End Sub

Private Sub TreeView1_Collapse(ByVal Node As ComctlLib.Node)
    If Not TreeView1.SelectedItem Is Nothing Then
        TreeView1_NodeClick TreeView1.SelectedItem
    End If
End Sub

Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
    If pvDelayLoad(m_oZip, TreeView1, Node) Then
        Set TreeView1.SelectedItem = Node
        TreeView1_NodeClick TreeView1.SelectedItem
    ElseIf Not TreeView1.SelectedItem Is Nothing Then
        TreeView1_NodeClick TreeView1.SelectedItem
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    pvDelayLoad m_oZip, TreeView1, Node
    RichTextBox1.TextRTF = ConcatCollection(pvEnumTags(Node))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '清除PEEK缓冲区安全阵列
    RollingHash 0, 0
End Sub
