VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'******************************************************************************
'** Class Name.....: ClsArrow
'** Description....:
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
Private Const m_ClassName = "ClsArrow"

Private mpicUp                    As Picture
Private mParent                   As PictureBox
Private mbIsDisplayed             As Boolean
Private mImgRect                  As RECT
Private miArrowType               As Integer
Private mlButtonHeight            As Long
Private mlState                   As Long
Private mbLastButtonDown          As Boolean

Private Const PIXELS_FROM_TOP     As Integer = 6
Private Const PIXELS_FROM_RIGHT   As Integer = 6
Private Const PIXELS_FROM_BOTTOM  As Integer = 6
Private Const PIXEL_WIDTH         As Integer = 16
Private Const ARROW_UP            As Integer = 1
Private Const ARROW_DOWN          As Integer = -1

'**************************
'****    Properties    ****
'**************************

Public Property Get objBitmap() As Object
  Set objBitmap = mpicUp
End Property

Public Property Set objBitmap(ByVal aobjBitmap As Object)
  Set mpicUp = aobjBitmap
End Property

Public Property Set ctlParent(ByVal actlParent As Control)
  Set mParent = actlParent
End Property

Public Property Get bIsDisplayed() As Boolean
  bIsDisplayed = mbIsDisplayed
End Property

Public Property Let iArrowType(ByVal aiArrowType As Integer)
  miArrowType = aiArrowType
End Property

Public Property Let lButtonHeight(ByVal alButtonHeight As Long)
  mlButtonHeight = alButtonHeight
End Property

'*******************************
'****    Public Methodes    ****
'*******************************

Public Sub Show(Optional ByVal alMenusAtTop As Long, Optional ByVal alMenusAtBottom As Long)
'******************************************************************************
'** SubRoutine.....: Show
'**
'** Description....: Draws the Arrows
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** alMenusAtTop        long       R     Number of menus at the top
'** alMenusAtBottom     Long       R     Number of menus at the bottom
'******************************************************************************
On Error GoTo ErrorSection
  
If Not mParent Is Nothing Then
    'Exit procedure if we've been called in internet explorer and
    'we're not visible yet.
    If Not mParent.Visible Then Exit Sub
    'Draw arrows
    With mImgRect
      .Left = mParent.ScaleWidth - PIXELS_FROM_RIGHT - PIXEL_WIDTH
      If miArrowType = ARROW_UP Then      'Up-Arrow
        If alMenusAtTop = 0 Then
          .Left = 0
          Exit Sub
        End If
        .Top = PIXELS_FROM_TOP + alMenusAtTop * mlButtonHeight
      Else                                'Down-Arrow
        .Top = mParent.ScaleHeight - PIXELS_FROM_BOTTOM - alMenusAtBottom * mlButtonHeight
      End If
      .Right = .Left + PIXEL_WIDTH
      .Bottom = .Top + PIXEL_WIDTH
      If .Left <> 0 Then
        mbIsDisplayed = True
        DrawBorder RAISED
      End If
    End With
End If
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
    ShowError Err.Number, Err.Description, "Show", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Public Sub Hide()
'******************************************************************************
'** SubRoutine.....: Hide
'**
'** Description....: Hids the Arrows
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** None
'******************************************************************************
On Error GoTo ErrorSection

  If mbIsDisplayed Then
    With mImgRect
      mParent.Line (.Left, .Top)-(.Right, .Bottom), BACKGROUND_COLOR, BF
    End With
    mbIsDisplayed = False
  End If
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
    ShowError Err.Number, Err.Description, "Hide", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Public Function bHitTest(ByVal aiMousePosition As Integer, _
                        ByVal alX As Long, _
                        ByVal alY As Long) As Boolean
'******************************************************************************
'** Function.......: bHitTest
'**
'** Description....: Tests which button has been hit and takes appropriate action
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** aiMousePosition     Integer     R    Position of Mouse
'** alX                 Long        R    X-Coordinate
'** alY                 Long        R    Y-Coordinate
'******************************************************************************
On Error GoTo ErrorSection
 
If mbIsDisplayed Then
    If PtInRect(mImgRect, alX, alY) <> 0 Then
        bHitTest = True
        Select Case aiMousePosition
            Case MOUSE_UP
                Select Case mlState
            Case SUNKEN
                DrawBorder RAISED
            Case Else
              ' nothing to do
        End Select
        mbLastButtonDown = False
        Case MOUSE_DOWN
            Select Case mlState
                Case SUNKEN
                    ' nothing to do
                Case Else
                    DrawBorder SUNKEN
            End Select
            mbLastButtonDown = True
        Case MOUSE_MOVE
            If mbLastButtonDown And mlState = RAISED Then
                DrawBorder SUNKEN
            End If
        End Select
    Else
        Select Case aiMousePosition
            Case MOUSE_UP
                mbLastButtonDown = False
            Case MOUSE_MOVE
                If mlState = SUNKEN Then DrawBorder RAISED
            Case MOUSE_DOWN
                If mlState = SUNKEN Then DrawBorder RAISED
                mbLastButtonDown = False
        End Select
    End If
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
    ShowError Err.Number, Err.Description, "bHitTest", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Function

Public Sub DrawBorder(ByVal aiDirection As Integer)
'******************************************************************************
'** SubRoutine.....: DrawBorder
'**
'** Description....: Draw an arrow
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** aiDirection         Integer    R     Direction of the arrow
'******************************************************************************
On Error GoTo ErrorSection
    
If mbIsDisplayed Then
    Select Case aiDirection
        Case RAISED
            With mImgRect
                mParent.PaintPicture mpicUp, .Left, .Top
            End With
            DrawEdge mParent.hDC, mImgRect, BDR_RAISED, BF_RECT
            mlState = RAISED
        Case SUNKEN
            With mImgRect
                mParent.PaintPicture mpicUp, .Left + 1, .Top + 1, .Right - .Left - 1, .Bottom - .Top - 1
            End With
            DrawEdge mParent.hDC, mImgRect, BDR_SUNKEN, BF_RECT
            mlState = SUNKEN
    End Select
End If
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

Public Sub Reset()
'******************************************************************************
'** SubRoutine.....: Reset
'**
'** Description....: Reset the arrow
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 16-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** None
'******************************************************************************
  mbLastButtonDown = False
  mbIsDisplayed = False
End Sub
