VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'******************************************************************************
'** Class Name.....: MenuItem
'** Description....: The Icon Pict and The Caption
'**
'**
'** Cie/Co ........: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'** Version........: 1.0.0.A
'**
'** Property             Data Type     Description
'** ------------------   ---------     --------------------------------------
'**
'** Method(Public)       Description
'** ------------------   --------------------------------------
'**
'** Event()              Description
'** ------------------   --------------------------------------
'**
'******************************************************************************
Private Const m_ClassName = "MenuItem"

Private picMenu As PictureBox
Private picCache As PictureBox
Private msCaption As String      ' caption of MenuItem
Private mlIndex As Long          ' index of icon on Menu (1 based)
Private picButton As StdPicture  ' icon picture
Private msCaptionX As Long
Private msCaptionY As Long
Private mlButtonHeight As Long
Private mbButtonDownOnMe As Boolean
Private msKey As String
Private msTag As String

Private Type BUTTON_STRUCT
    RECT     As RECT
    State    As Long
    OnScreen As Boolean
End Type

Private mButtonStruct As BUTTON_STRUCT
Private mHitStruct As RECT
Private m3DStruct As RECT

Const HITTEXT_EXTRA_PIXELS = 4
Private Const CLIPPING_NO = True
Private Const CLIPPING_YES = False


'**************************
'****    Properties    ****
'**************************

Public Property Get Caption() As String
    Caption = msCaption
End Property

Public Property Let Caption(ByVal sNewValue As String)
    msCaption = sNewValue
End Property

Public Property Get Index() As Long
    Index = mlIndex
End Property

Public Property Let Index(ByVal lNewValue As Long)
    mlIndex = lNewValue
End Property

Public Property Get Button() As StdPicture
    Set Button = picButton
End Property

Public Property Set Button(ByVal vNewValue As StdPicture)
    Set picButton = vNewValue
End Property

Public Property Get Left() As Long
    Left = mButtonStruct.RECT.Left
End Property

Public Property Let Left(ByVal lNewValue As Long)
    With mButtonStruct.RECT
        .Left = lNewValue
        .Right = lNewValue + SIZE_ICON
    End With
End Property

Public Property Get Top() As Long
    Top = mButtonStruct.RECT.Top
End Property

Public Property Let Top(ByVal lNewValue As Long)
    With mButtonStruct.RECT
        .Top = lNewValue
        .Bottom = lNewValue + SIZE_ICON
    End With
End Property

Public Property Get Right() As Long
    Right = mButtonStruct.RECT.Right
End Property

Public Property Get Bottom() As Long
    Bottom = mButtonStruct.RECT.Bottom
End Property

Public Property Get State() As Long
    State = mButtonStruct.State
End Property

Public Property Let State(ByVal lNewValue As Long)
    mButtonStruct.State = lNewValue
End Property

Public Property Get CaptionX() As Long
    CaptionX = msCaptionX
End Property

Public Property Let CaptionX(ByVal lNewValue As Long)
    msCaptionX = lNewValue
End Property

Public Property Get CaptionY() As Long
    CaptionY = msCaptionY
End Property

Public Property Let CaptionY(ByVal lNewValue As Long)
    msCaptionY = lNewValue
End Property

Public Property Set Parent(ByVal picNewValue As Control)
    Set picMenu = picNewValue
End Property

Public Property Set Cache(ByVal oNewValue As Object)
    Set picCache = oNewValue
End Property

Public Property Let ButtonHeight(ByVal lNewValue As Long)
    mlButtonHeight = lNewValue
End Property

Public Property Get Key() As String
    Key = msKey
End Property

Public Property Let Key(ByVal sNewValue As String)
    msKey = sNewValue
End Property

Public Property Get Tag() As String
    Tag = msTag
End Property

Public Property Let Tag(ByVal sNewValue As String)
    msTag = sNewValue
End Property

'*******************************
'****    Public Methodes    ****
'*******************************

Public Function HitTest(ByVal iMousePosition As Integer, _
                            ByVal x As Long, _
                            ByVal y As Long) As Boolean
'******************************************************************************
'** Function.......: HitTest
'**
'** Description....: Set Border effect on Mouse
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** iMousePosition      Integer    R     Mouse Up . down
'** x                   Long       R     X-Coord Mouse
'** y                   Long       R     Y-Coord Mouse
'** HitTest             Boolean    W     TRUE if MOUSE DOWN
'******************************************************************************
On Error GoTo ErrorSection
Dim bVal As Boolean
'** Default
bVal = False
'** don't bother if it is not on screen
If Not mButtonStruct.OnScreen Then
    GoTo ExitPoint
End If
    
If PtInRect(mHitStruct, x, y) Then
    bVal = True
    Select Case iMousePosition
        Case MOUSE_UP
            Select Case mButtonStruct.State
                Case NONE, SUNKEN
                    DrawBorder RAISED
                Case Else
                    ' nothing to do
            End Select
            mbButtonDownOnMe = False
        Case MOUSE_DOWN
            Select Case mButtonStruct.State
                Case SUNKEN
                    ' nothing to do - it's already drawn
                Case Else
                    DrawBorder SUNKEN
                    mbButtonDownOnMe = True
            End Select
        Case MOUSE_MOVE
            Select Case mButtonStruct.State
                Case RAISED
                    ' nothing to do - it's already drawn
                Case NONE
                    ' if the mouse went down on me, moved off me
                    ' and now returns and no mouse up yet, draw
                    ' me as down
                    If Not mbButtonDownOnMe Then
                        DrawBorder RAISED
                    Else
                        DrawBorder SUNKEN
                    End If
                Case SUNKEN
                    ' leave it that way
            End Select
    End Select
Else
    '** there is no hit
    bVal = False
    If iMousePosition <> MOUSE_MOVE Then
        mbButtonDownOnMe = False
    End If
    '** if any border is currently drawn, remove it
    If mButtonStruct.State <> NONE Then
        DrawBorder NONE
    End If
End If
    
'********************
'Exit Point
'********************
ExitPoint:
HitTest = bVal
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "HitTest", m_ClassName, vbLogEventTypeError
End Select
bVal = False
Resume ExitPoint
    
End Function

Public Sub DrawBorder(ByVal iDirection As Integer)
'******************************************************************************
'** SubRoutine.....: DrawBorder
'**
'** Description....: Draw the Border Effect
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** iDirection          Integer    R     Mouse Over Down or not
'******************************************************************************
On Error GoTo ErrorSection
     
If Not mButtonStruct.OnScreen Then
    Exit Sub
End If
    
State = iDirection
Select Case iDirection
    Case RAISED
        DrawEdge picMenu.hDC, m3DStruct, BDR_RAISEDOUTER, BF_RECT
    Case SUNKEN
        DrawEdge picMenu.hDC, m3DStruct, BDR_SUNKENINNER, BF_RECT
    Case NONE
        With m3DStruct
            picMenu.Line (.Left, .Top)-(.Right - 1, .Bottom - 1), BACKGROUND_COLOR, B
        End With
End Select
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
    ShowError Err.Number, Err.Description, "DrawBorder", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Public Function PaintButton(ByVal lTopMenuItemDisplayed As Long, _
                        ByVal lIconStart As Long, _
                        ByVal lMenuCur As Long, _
                        ByVal lClipY As Long) As Boolean
'******************************************************************************
'** SubRoutine.....: PaintButton
'**
'** Description....:
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 17-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**
'******************************************************************************
On Error GoTo ErrorSection

Dim lCenter As Long
Dim lLeft As Long
Dim lTop As Long
Dim lRight As Long
Dim lBottom As Long
Dim lResult As Long
Dim lHeight As Long
Dim bClipping As Boolean
Dim lPositionFromTop As Long
Dim RgnRect As RECT
Dim hRgn As Long
Dim hSavedDC As Long
Dim bVal As Boolean

bVal = CLIPPING_NO

If mlIndex < lTopMenuItemDisplayed Then
    mButtonStruct.OnScreen = False
    GoTo ExitPoint
End If
    
'** position the image
lPositionFromTop = mlIndex - lTopMenuItemDisplayed + 1
Top = (lPositionFromTop * 2 * SIZE_ICON) - SIZE_ICON + ((lPositionFromTop + 1 = 1) * 4) + (lMenuCur - 1) * mlButtonHeight
With picMenu
    lCenter = .ScaleWidth \ 2
    Left = lCenter - (SIZE_ICON \ 2)
End With
    
With mButtonStruct
    lLeft = .RECT.Left
    lTop = .RECT.Top
    lRight = .RECT.Right
    lBottom = .RECT.Bottom
    '** see if it will fit in the control's viewing area
    If lTop > lClipY Then
        .OnScreen = False
        bVal = CLIPPING_YES
        GoTo ExitPoint
    End If
    If lBottom > lClipY Then
        bClipping = True
        lBottom = lClipY
    End If
    .OnScreen = True
End With
    
'** position the menu caption
CaptionX = lCenter - (CLng(picMenu.TextWidth(Caption())) \ 2)
CaptionY = lTop + SIZE_ICON + 4

'** calculate the hittest structure
With mHitStruct
    .Left = lLeft - HITTEXT_EXTRA_PIXELS - 2
    .Top = lTop - HITTEXT_EXTRA_PIXELS - 2
    .Right = lRight + HITTEXT_EXTRA_PIXELS + 2
    '** hittest includes the caption below the icon
    .Bottom = lBottom + picMenu.TextHeight(Caption()) + 5
    If bClipping Then
        .Bottom = lBottom
    End If
End With
        
'** calculate the 3d structure
With m3DStruct
    .Left = lLeft - 2
    .Top = lTop - 2
    .Right = lRight + 2
    '** hittest includes the caption below the icon
    If Not bClipping Then
        .Bottom = lBottom + 2
    Else
        .Bottom = lBottom
    End If
End With

With mButtonStruct.RECT
    If Not bClipping Then
        lHeight = SIZE_ICON
    Else
        lHeight = lBottom - lTop
    End If
    BitBlt picMenu.hDC, .Left, .Top, SIZE_ICON, lHeight, _
           picCache.hDC, 0, mlButtonHeight * 2 + (lIconStart + lPositionFromTop) * SIZE_ICON, SRCCOPY
End With

'** position the caption
If Not bClipping Then
    With picMenu
        .CurrentX = msCaptionX
        .CurrentY = msCaptionY
        .ForeColor = ITEMFORCOLOR
        If .CurrentY + .TextHeight(msCaption) < lClipY Then
            picMenu.Print msCaption
            bVal = CLIPPING_NO
        Else
            '** set the region
            With picMenu
                RgnRect.Left = 0
                RgnRect.Top = msCaptionY
                RgnRect.Right = .Width
                RgnRect.Bottom = lClipY
                '** save the original DC
                hSavedDC = SaveDC(.hDC)
                '** create a region for the text
                hRgn = CreateRectRgnIndirect(RgnRect)
                '** set clipping
                SelectClipRgn .hDC, hRgn
                '** print the caption
                picMenu.Print msCaption
                '** delete the object
                DeleteObject hRgn
                '** restore the original DC
                RestoreDC .hDC, hSavedDC
                bVal = CLIPPING_YES
            End With
        End If
    End With
Else
    bVal = CLIPPING_YES
End If

'********************
'Exit Point
'********************
ExitPoint:
PaintButton = bVal
Exit Function
'********************
'Error Section
'********************
ErrorSection:
Select Case Err.Number
    Case Else
    ShowError Err.Number, Err.Description, "PaintButton", m_ClassName, vbLogEventTypeError
End Select
bVal = False
Resume ExitPoint

End Function

