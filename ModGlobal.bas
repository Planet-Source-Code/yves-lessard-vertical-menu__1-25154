Attribute VB_Name = "ModGlobal"
Option Explicit
'******************************************************************************
'** Module.........: ModGlobal
'**
'** Description....: Global Constant & API
'**                  Also Error Handling
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 17-Jul-2001.
'**
'** Modifications..:
'** Version........: 1.0.0.A
'**
'******************************************************************************
Private Const m_ClassName = "ModGlobal"

Public BACKGROUND_COLOR As Long
Public SIZE_ICON As Integer
Public MENUFORCOLOR As OLE_COLOR
Public ITEMFORCOLOR As OLE_COLOR

Public Const RAISED = 1
Public Const SUNKEN = -1
Public Const NONE = 0
Public Const BUTTON_NONE = 0
Public Const BUTTON_UP = 1
Public Const BUTTON_DOWN = 2
Public Const MOUSE_UP = 1
Public Const MOUSE_DOWN = -1
Public Const MOUSE_MOVE = 0

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP _
             Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_RIGHT)

Public Const SRCCOPY = &HCC0020

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

Public Type POINTAPI
    x   As Long
    y   As Long
End Type

Public Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

'** API
Public Declare Function PtInRect Lib "user32" (RECT As RECT, ByVal lPtX As Long, ByVal lPtY As Long) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hMF As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal SavedDC As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean


'*********************************
'****    Error(s) Handling    ****
'*********************************
Public Sub ShowError(ErrorNumber As Long, ErrorMsg As String _
                      , ErrorModule As String, ErrorForm As String _
                     , LogEventType As Long, Optional ErrorInfo As Variant)
'******************************************************************************
'** Module.........: ShowError
'** Description....: This routine is used to show the current
'**                  error Message and LOG the error to a file.
'**
'** Author, date...: Yves Lessard , 16-Jul-2001.
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

