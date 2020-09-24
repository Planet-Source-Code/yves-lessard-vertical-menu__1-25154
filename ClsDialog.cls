VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'******************************************************************************
'** Module.........: ClsDialog
'** Description....: Some Dialogs utilities
'**                  Private MonDlg As ClsDialog
'**                  Set MonDlg = New ClsDialog
'**                  Set MonDlg.Form = Me
'**
'** Author, date...: Yves Lessard , 17-Jul-2000.
'**
'** Version........: 1.0.0.A
'** Modifications..:
'** 1.0.1 - (SevySoft), dd-mmm-yyyy
'**
'** Property             Data Type     Description
'** ------------------   ---------     --------------------------------------
'** Bold                 Boolean       Read/Write from Font
'** FontName             String        Read/Write Font
'** FontSize             Long          Read/Write Font
'** Italic               Boolean       Read/Write Font
'** StrikeThru           Boolean       Read/Write Font
'** Underline            Boolean       Read/Write Font
'** FileName             String        From ShowOpen & ShowSave
'** Filter               String        Filter Option "Text files|*.doc|All files|*.*"
'**                                    For Open & Save Dialog
'** Form                 Form          Assign the form in the LoadEvent
'** InitDir              String        The Initial Dir for Save & Open
'**
'** Method(Public)       Description
'** ------------------   --------------------------------------
'** FormatDrive          Show dialog to format Floppy
'** Reboot               Show reboot dialog
'** ShowColor            The Color Dialog
'** ShowFont             The Font Dialog
'** ShowOpen             The Open File dialog
'** ShowSave             The Save File Dialog
'** ShowShutDown         The ShutDown Dialog
'**
'******************************************************************************
Private Const m_ClassName = "ClsDialog"
Private mhOwner As Long
Private Const Restart_Reboot = &H2
Private Const cdlgOpen = 1
Private Const LOGPIXELSY = 90
Private Const FW_BOLD = 700
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100

Private mvarInitDir As String
Private mFileName As String
Private mvarFilter As String
Private mFileTitle As String
Private mDefaultExt As String
Private mFilterIndex As Long
Private mDialogTitle As String
Private mvarRGBcolor As Long
Private mvarFontName As String
Private mvarFontSize As Long
Private mvarItalic As Boolean
Private mvarBold As Boolean
Private mvarTheForm As Form
Private mUnderline As Boolean
Private mStrikethru As Boolean

Private Type OPENFILENAME 'Open & Save Dialog
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSECOLOR 'Color Dialog
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type ChooseFont 'Font Dialog
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Const LF_FACESIZE = 32 'Font Dialog
Private Type LOGFONT 'Font Dialog
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Enum FontsConstants
    cdlCFScreenFonts = 1
    cdlCFPrinterFonts = 2
    cdlCFBoth = 3
    cdlCFHelpButton = 4
    cdlCFEffects = 256
    cdlCFApply = 512
    cdlCFANSIOnly = 1024
    cdlCFNoVectorFonts = 2048
    cdlCFNoSimulations = 4096
    cdlCFLimitSize = 8192
    cdlCFFixedPitchOnly = 16384
    cdlCFWYSIWYG = 32768
    cdlCFForceFontExist = 65536
    cdlCFScalableOnly = 131072
    cdlCFTTOnly = 262144
    cdlCFNoFaceSel = 524288
    cdlCFNoStyleSel = 1048576
    cdlCFNoSizeSel = 2097152
End Enum

Const MAX_PATH = 260 'General
Private OFN As OPENFILENAME ' Open & Save Dialog
'Color Dialog
Private Const cdlCCFullOpen = &H2
Private Const cdlCCHelpButton = &H8
Private Const cdlCCPreventFullOpen = &H4
Private Const cdlCCRGBInit = &H1
'*************************
'***   API Functions   ***
'*************************
Private Declare Function SHShutDownDialog Lib "Shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Declare Function SHRestartSystem Lib "Shell32" Alias "#59" (ByVal hOwner As Long, ByVal sPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHFormatDrive Lib "Shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetDC Lib "user32" (ByVal hOwner As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (p1 As Any, p2 As Any) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hDC As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long



'**************************
'****    Properties    ****
'**************************

Public Property Get StrikeThru() As Boolean
   StrikeThru = mStrikethru
End Property

Public Property Let StrikeThru(bStrikethru As Boolean)
   mStrikethru = bStrikethru
End Property

Public Property Get Underline() As Boolean
'******************************************************************************
'** Description....: Return Underline Font status
'** Author, date...: Yves Lessard  07-Aug-2000.
'******************************************************************************
   Underline = mUnderline
End Property

Public Property Let Underline(bUnderline As Boolean)
'******************************************************************************
'** Description....: Assign Underline font
'** Author, date...: Yves Lessard  07-Aug-2000.
'******************************************************************************
   mUnderline = bUnderline
End Property

Public Property Set Form(ByVal vData As Form)
'******************************************************************************
'** Description....: Assign the Form parent
'** Author, date...: Yves Lessard  07-Aug-2000.
'******************************************************************************
    Set mvarTheForm = vData
    mhOwner = Form.hWnd
    '** Init Font to the same as Parent Form
    InitDefaultFont
End Property

Public Property Get Form() As Form
'******************************************************************************
'** Description....: get the Form parent
'** Author, date...: Yves Lessard  07-Aug-2000.
'******************************************************************************
    Set Form = mvarTheForm
End Property

Public Property Let Italic(ByVal vData As Boolean)
'******************************************************************************
'** Description....: Assign Italic
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    mvarItalic = vData
End Property

Public Property Get Italic() As Boolean
'******************************************************************************
'** Description....: Retourn Italic
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    Italic = mvarItalic
End Property

Public Property Let Bold(ByVal vData As Boolean)
'******************************************************************************
'** Description....: Assign Bold for Font
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    mvarBold = vData
End Property

Public Property Get Bold() As Boolean
'******************************************************************************
'** Description....: Retourn bold value
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    Bold = mvarBold
End Property

Public Property Let FontSize(ByVal vData As Long)
'******************************************************************************
'** Description....: Assign FontSize
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    mvarFontSize = vData
End Property

Public Property Get FontSize() As Long
'******************************************************************************
'** Description....: Get FontSize
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    FontSize = mvarFontSize
End Property

Public Property Let FontName(ByVal vData As String)
'******************************************************************************
'** Description....: Assign Font Name
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    mvarFontName = vData
End Property

Public Property Get FontName() As String
'******************************************************************************
'** Description....: Get the FontName
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    FontName = mvarFontName
End Property

Public Property Get RGBcolor() As Long
'******************************************************************************
'** Description....: Get the RGBcolor
'** Author, date...: Yves Lessard  25-Jul-2000.
'******************************************************************************
    RGBcolor = mvarRGBcolor
End Property

Public Property Let Filter(ByVal vData As String)
'******************************************************************************
'** Description....: Filter Option for ShowOpen
'** Example: MonDlg.Filter = "Text files|*.txt|All files|*.*"
'** Author, date...: Yves Lessard  18-Jul-2000.
'******************************************************************************
    mvarFilter = vData
End Property

Public Property Let InitDir(ByVal vData As String)
'******************************************************************************
'** Description....: Directory to look ShowOpen
'** Author, date...: Yves Lessard  18-Jul-2000.
'******************************************************************************
    mvarInitDir = vData
End Property


Public Property Get InitDir() As String
'******************************************************************************
'** Description....: The directory return by ShowOpen
'** Author, date...: Yves Lessard  18-Jul-2000.
'******************************************************************************
    InitDir = mvarInitDir
End Property

Public Property Let FileName(ByVal vData As String)
'******************************************************************************
'** Description....: FileName for ShowOpen
'** Author, date...: Yves Lessard  18-Jul-2000.
'******************************************************************************
    mFileName = vData
End Property

Public Property Get FileName() As String
'******************************************************************************
'** Description....: FileName return by ShowOpen
'** Author, date...: Yves Lessard  18-Jul-2000.
'******************************************************************************
    FileName = mFileName
End Property


'************************
'****    Methodes    ****
'************************

Public Sub ShowShutDown()
'******************************************************************************
'** SubRoutine.....: ShowShutDown
'**
'** Description....: Shut Down Dialog
'**
'** Author, date...: Yves Lessard , 17-Jul-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'******************************************************************************
On Error GoTo ErrorSection

   SHShutDownDialog mhOwner
'********************
'Exit Point
'********************
ExitPoint:
Exit Sub
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "ShowShutDown", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint
   
End Sub


Private Sub InitOFN()
'******************************************************************************
'** SubRoutine.....: InitOFN
'**
'** Description....: Init structure for ShowOpen
'**
'** Author, date...: Yves Lessard , 18-Jul-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'******************************************************************************
On Error GoTo ErrorSection

Dim sTemp As String, i As Integer
Dim uFlag As Long
If Me.InitDir = "" Then
    Me.InitDir = App.Path & "\"
End If

uFlag = cdlgOpen And (&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
With OFN
    .lStructSize = Len(OFN)
    .hwndOwner = mhOwner
    .flags = uFlag
    .lpstrDefExt = mDefaultExt
    sTemp = Me.InitDir
    If sTemp = "" Then sTemp = App.Path
    .lpstrInitialDir = sTemp
    sTemp = mFileName
    .lpstrFile = sTemp & String$(255 - Len(sTemp), 0)
    .nMaxFile = 255
    .lpstrFileTitle = String$(255, 0)
    .nMaxFileTitle = 255
    sTemp = mvarFilter
    For i = 1 To Len(sTemp)
        If Mid$(sTemp, i, 1) = "|" Then
            Mid$(sTemp, i, 1) = vbNullChar
        End If
    Next
    sTemp = sTemp & String$(2, 0)
    .lpstrFilter = sTemp
    .nFilterIndex = mFilterIndex
    .lpstrTitle = mDialogTitle
    .hInstance = App.hInstance
End With
'********************
'Exit Point
'********************
ExitPoint:
Exit Sub
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "InitOFN", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Public Sub Reboot()
'******************************************************************************
'** SubRoutine.....: Reboot
'**
'** Description....: Show Reboot Dialog
'**
'** Author, date...: Yves Lessard , 17-Jul-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'******************************************************************************
On Error GoTo ErrorSection
Dim uFlag As Long
uFlag = Restart_Reboot And (&H0 Or &H1 Or &H2 Or &H4)
SHRestartSystem mhOwner, "", uFlag

'********************
'Exit Point
'********************
ExitPoint:
Exit Sub
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "Reboot", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Public Function ShowSave() As Boolean
'******************************************************************************
'** SubRoutine.....: ShowSave
'**
'** Description....: The Show Save Dialog file
'**
'** Author, date...: Yves Lessard , 25-Jul-2000.
'**
'** Example..:     With MonDlg
'**                    .Filter = "Text files|*.txt|All files|*.*"
'**                End With
'**                If MonDlg.ShowSave Then
'**                    MsgBox MonDlg.FileName
'**                End If
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**  ShowSave           Boolean     W    TRUE -> The file saved
'******************************************************************************
On Error GoTo ErrorSection
Dim bFlag As Boolean
Dim iDelim As Integer
Dim RetValue As Long

'** Default File saved
bFlag = True
InitOFN
RetValue = GetSaveFileName(OFN)

If RetValue > 0 Then
    iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
    If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
    iDelim = InStr(OFN.lpstrFile, vbNullChar)
    If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
Else
    Me.FileName = ""
    bFlag = False
End If
'********************
'Exit Point
'********************
ExitPoint:
ShowSave = bFlag
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "ShowSave", m_ClassName, vbLogEventTypeError
End Select
Me.FileName = ""
bFlag = False
Resume ExitPoint

End Function

Public Function ShowOpen() As Boolean
'******************************************************************************
'** SubRoutine.....: ShowOpen
'**
'** Description....: The Show Open Dialog file
'**
'** Author, date...: Yves Lessard , 18-Jul-2000.
'**
'** Example..:     With MonDlg
'**                   .Filter = "Text files|*.doc|All files|*.*"
'**                   .FileName = "c:\autoexec.bat"
'**                End With
'**                If MonDlg.ShowOpen Then
'**                    MsgBox MonDlg.FileName
'**                End If
'**
'** if me.FileName ="" means no file was choose or cancel action
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**  ShowOpen           Boolean     W    TRUE -> File selection
'******************************************************************************
On Error GoTo ErrorSection
Dim bFlag As Boolean
Dim iDelim As Integer
Dim RetValue As Long

'** Default we have a file
bFlag = True
  InitOFN
  RetValue = GetOpenFileName(OFN)
  If RetValue > 0 Then
     iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
     If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
     iDelim = InStr(OFN.lpstrFile, vbNullChar)
     If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
  Else
     Me.FileName = ""
     bFlag = False
  End If
'********************
'Exit Point
'********************
ExitPoint:
ShowOpen = bFlag
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "ShowOpen", m_ClassName, vbLogEventTypeError
End Select
Me.FileName = ""
bFlag = False
Resume ExitPoint
  
End Function

Public Function ShowColor() As Boolean
'******************************************************************************
'** SubRoutine.....: ShowColor
'**
'** Description....: Show the Color Dialog
'**
'** Author, date...: Yves Lessard , 25-Jul-2000.
'**
'**   Exemple:      If MonDlg.ShowColor Then
'**                     Me.BackColor = MonDlg.RGBcolor
'**                 End If
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** ShowColor           Boolean     W    True -> Color choosen else Cancel
'******************************************************************************
On Error GoTo ErrorSection

Dim CC As CHOOSECOLOR
Dim CustomColors() As Byte
Dim uFlag As Long
ReDim CustomColors(0 To 16 * 4 - 1) As Byte
Dim i As Integer
Dim bFlag As Boolean
Dim RetValue As Long

bFlag = True
For i = LBound(CustomColors) To UBound(CustomColors)
    CustomColors(i) = 255 ' white
Next

uFlag = cdlCCFullOpen And (&H1 Or &H2 Or &H4 Or &H8)
With CC
    .lStructSize = Len(CC)
    .hwndOwner = mhOwner
    .hInstance = App.hInstance
    .lpCustColors = StrConv(CustomColors, vbUnicode)
    .flags = uFlag
    .RGBResult = mvarRGBcolor
    RetValue = ChooseColorAPI(CC)
    If RetValue = 0 Then
        bFlag = False
    Else
        CustomColors = StrConv(.lpCustColors, vbFromUnicode)
        mvarRGBcolor = .RGBResult
        bFlag = True
    End If
End With
'********************
'Exit Point
'********************
ExitPoint:
ShowColor = bFlag
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "ShowColor", m_ClassName, vbLogEventTypeError
End Select
bFlag = False
Resume ExitPoint
  
End Function

Public Function FormatDrive(Optional ByVal iDrive As Long, Optional ByVal iCapacity As Long, Optional ByVal iFormatType As Long) As Boolean
'******************************************************************************
'** Function.......: FormatDrive
'**
'** Description....: Format a Drive Default is Floppy Drive
'**                  Always show up dialog
'**
'** Author, date...: Yves Lessard , 17-Jul-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'******************************************************************************
On Error GoTo ErrorSection
Dim result As Long

result = SHFormatDrive(mhOwner, iDrive, iCapacity, iFormatType)
If result = -1 Then
    FormatDrive = False
Else
    FormatDrive = True
End If
'********************
'Exit Point
'********************
ExitPoint:
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "FormatDrive", m_ClassName, vbLogEventTypeError
End Select
FormatDrive = False
Resume ExitPoint

End Function

Public Sub InitDefaultFont()
'******************************************************************************
'** SubRoutine.....: InitDefaultFont
'**
'** Description....: Get the same font and so on as the Parent Form
'**
'** Author, date...: Yves Lessard , 07-Aug-2000.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'******************************************************************************
On Error GoTo ErrorSection

With Me
    .FontName = Form.FontName
    .FontSize = Form.FontSize
    .Bold = Form.FontBold
    .Italic = Form.FontItalic
End With
'********************
'Exit Point
'********************
ExitPoint:
Exit Sub
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "InitDefaultFont", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint


End Sub

Public Function ShowFont() As Boolean
'******************************************************************************
'** SubRoutine.....: ShowFont
'**
'** Description....: Display the Font dialog
'**
'** Author, date...: Yves Lessard , 07-Aug-2000.
'**
'** Example..:      If MonDlg.ShowFont Then
'**                    Command1(11).FontBold = MonDlg.Bold
'**                    Command1(11).FontName = MonDlg.FontName
'**                    Command1(11).FontSize = MonDlg.FontSize
'**                 End If
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**  ShowFont           Boolean     W    TRUE -> Something change
'******************************************************************************
On Error GoTo ErrorSection

Dim bFlag As Boolean
Dim CF As ChooseFont
Dim LF As LOGFONT
Dim TempByteArray() As Byte
Dim ByteArrayLimit As Long
Dim OldhDC As Long
Dim FontToUse As Long
Dim tbuf As String * 80
Dim x As Long
Dim uFlag As Long
Dim RetValue As Long

'** Default something change
bFlag = True
uFlag = 4 And (&H1 Or &H2 Or &H3 Or &H4 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H10000 Or &H20000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
TempByteArray = StrConv(FontName & vbNullChar, vbFromUnicode)
ByteArrayLimit = UBound(TempByteArray)
  
With LF
    For x = 0 To ByteArrayLimit
        .lfFaceName(x) = TempByteArray(x)
    Next
    .lfHeight = FontSize / 72 * GetDeviceCaps(GetDC(mhOwner), LOGPIXELSY)
    .lfItalic = Italic * -1
    .lfUnderline = Underline * -1
    .lfStrikeOut = StrikeThru * -1
    If Bold Then .lfWeight = FW_BOLD
End With

With CF
    .lStructSize = Len(CF)
    .hwndOwner = mhOwner
    .hDC = GetDC(mhOwner)
    .lpLogFont = lstrcpy(LF, LF)
    If Not uFlag Then
        .flags = cdlCFScreenFonts
    Else
        .flags = uFlag Or cdlCFWYSIWYG
    End If
    .flags = .flags Or cdlCFEffects Or CF_INITTOLOGFONTSTRUCT
    .rgbColors = vbBlack
    .lCustData = 0
    .lpfnHook = 0
    .lpTemplateName = 0
    .hInstance = 0
    .lpszStyle = 0
    .nFontType = SCREEN_FONTTYPE
    .nSizeMin = 0
    .nSizeMax = 0
    .iPointSize = FontSize * 10
End With

RetValue = ChooseFont(CF)
If RetValue = 0 Then
    '** Nothing Change
    bFlag = False
Else
    With LF
        Italic = .lfItalic * -1
        Underline = .lfUnderline * -1
        StrikeThru = .lfStrikeOut * -1
    End With
    With CF
        FontSize = .iPointSize \ 10
        If .nFontType And BOLD_FONTTYPE Then
            Bold = True
        Else
            Bold = False
        End If
    End With
    FontToUse = CreateFontIndirect(LF)
    If FontToUse <> 0 Then
        OldhDC = SelectObject(CF.hDC, FontToUse)
        RetValue = GetTextFace(CF.hDC, 79, tbuf)
        FontName = Mid$(tbuf, 1, RetValue)
    Else
        bFlag = False
    End If
End If
    
'********************
'Exit Point
'********************
ExitPoint:
ShowFont = bFlag
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "ShowFont", m_ClassName, vbLogEventTypeError
End Select
bFlag = False
Resume ExitPoint
       
End Function

'*******************************
'****    Errors Handling    ****
'*******************************

Private Sub ShowError(ErrorNumber As Long, ErrorMsg As String _
                      , ErrorModule As String, ErrorForm As String _
                     , LogEventType As Long, Optional ErrorInfo As Variant)
'******************************************************************************
'** Module.........: ShowError
'** Description....: This routine is used to show the current
'**                  error Message and LOG the error to a file.
'**
'** Author, date...: Yves Lessard , 17-Jul-2000.
'**
'** Name                Type     Acces   Description
'** ------------------  -------  ------  --------------------------------------
'**  ErrorNumber         Long      R      Error Number
'**  ErrorMsg            String    R      Error Message
'**  ErrorModule         String    R      Module name where the error occured
'**  ErrorForm           String    R      Form Name where the error occured
'**  LogEventType        Long      R      Log event type (vbLogEventTypeError ,
'**                                       vbLogEventTypeWarning , vbLogEventTypeInformation)
'**  ErrorInfo           Variant   R      Additional error Information to Display
'**
'******************************************************************************
On Error GoTo ErrorSection
Dim ErrorTitle As String
Dim ErrorMessage As String

ErrorTitle = "ERROR - " & ErrorNumber & " - " & ErrorModule & " - " & ErrorForm
ErrorMessage = "ERROR  " & ErrorNumber & " - " & ErrorMsg

If Not IsMissing(ErrorInfo) Then
    ErrorMessage = ErrorMessage & vbCrLf & ErrorInfo
End If

MsgBox ErrorMessage, vbOKOnly + vbExclamation, ErrorTitle

App.LogEvent ErrorTitle & ": " & ErrorMessage, LogEventType

ExitPoint:
Exit Sub

ErrorSection:
Resume ExitPoint

End Sub

