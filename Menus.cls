VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'******************************************************************************
'** Class Name.....: Menus
'** Description....: Handle Menu Bouton
'**
'**
'** Cie/Co ........: SevySoft
'** Author, date...: Yves Lessard , 17-Jul-2001.
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
Private Const m_ClassName = "Menus"

Private picMenu As PictureBox
Private picCache As PictureBox
Private colMenus As New Collection
Private mlButtonHeight As Long
Private mlMenuPrev As Long
Private mlMenuCur As Long
Private mbNumberOfMenusChanged As Boolean
Private Const PIXELS_PER_BITBLT = 1
Private Const TYPE_UP = 1
Private Const TYPE_DOWN = -1
Private Const PIXELS_PER_BITBLT_S = 1

Public Function Add(ByVal sCaption As String, _
                ByVal lIndex As Long, _
                ByVal picMenu As Object) As VMenu
'******************************************************************************
'** Function.......: Add
'**
'** Description....: add a new Menu to the collection
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 18-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'**
'******************************************************************************
On Error GoTo ErrorSection

Dim newMenu As New VMenu
    
With newMenu
    .Caption = sCaption
    .Index = lIndex
    Set .Control = picMenu
    .ButtonHeight = mlButtonHeight
End With
        
If colMenus.Count = 0 Then
    colMenus.Add newMenu
ElseIf lIndex = colMenus.Count + 1 Then
    colMenus.Add newMenu
ElseIf lIndex = 1 Then
    colMenus.Add newMenu, , 1
Else
    colMenus.Add newMenu, , , lIndex - 1
End If

Set Add = newMenu
Set newMenu = Nothing
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
    ShowError Err.Number, Err.Description, "Add", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint
End Function

Public Sub Delete(ByVal lIndex As Long)
'******************************************************************************
'** SubRoutine.....: Delete
'**
'** Description....: delete the Menu from the collection
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 18-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** lIndex              Long        R    The Index to remove
'******************************************************************************
    On Error Resume Next
    colMenus.Remove lIndex
End Sub

Public Property Get Item(lIndex As Variant) As VMenu
'******************************************************************************
'** Description....: return the object of the Menu in the collection
'** Author, date...: Yves Lessard  18-Jul-2001.
'******************************************************************************
    If lIndex > 0 Then
        Set Item = colMenus(lIndex)
    End If
End Property

Public Property Get Count() As Long
    Count = colMenus.Count
End Property

Public Property Get Caption(lIndex As Long) As String
    Caption = colMenus(lIndex).Caption
End Property

Public Property Let Caption(lIndex As Long, sNewValue As String)
    colMenus(lIndex).Caption = sNewValue
End Property

Public Property Get ButtonHeight() As Long
    ButtonHeight = mlButtonHeight
End Property

Public Property Let ButtonHeight(ByVal lNewValue As Long)
    mlButtonHeight = lNewValue
End Property

Public Property Set Menu(oNewValue As PictureBox)
    Set picMenu = oNewValue
End Property

Public Property Set Cache(oNewValue As PictureBox)
    Set picCache = oNewValue
End Property

Public Property Let MenuCur(lNewValue As Long)
    mlMenuCur = lNewValue
End Property

' Procedure: Paint
' This is the main procedure that paints our control
' It handles repaints as well as well as changes of the
' current menu
' Since we can move several menus at once, the code for
' this is done here in the collection of menus rather then
' the the menu class itself.  However, the painting of the
' MenuItems is done in the MenuItem class itself.
Public Sub Paint()
    If mlMenuPrev = 0 Then               ' first time paint
        mlMenuPrev = mlMenuCur
    End If
    If mlMenuPrev = mlMenuCur Then
        Repaint
    ElseIf mlMenuPrev < mlMenuCur Then    ' user selected a menu after the previously selected menu
        ReselectDown
    Else                                ' user selected a menu before the previously selected menu
        ReselectUp
    End If
    
    DrawIcons
    SetMenuButtonsHotSpot
    mlMenuPrev = mlMenuCur                ' save this menu as the next previous menu
End Sub

Private Sub Repaint()
'******************************************************************************
'** SubRoutine.....: Repaint
'**
'** Description....: repaint the menu as is
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 18-Jul-2001.
'**
'** Modifications..:
'******************************************************************************
On Error GoTo ErrorSection

Dim l As Long
Dim lMax As Long
Dim hDestDC As Long
Dim hSrcDC As Long
Dim sCaption As String
Dim lWidth As Long
Dim lHeight As Long

' setup variables
lMax = colMenus.Count
With picMenu
    ' if we just changed the number of menus then
    ' we need to clear the control first
    If mbNumberOfMenusChanged Then
        .Cls
        mbNumberOfMenusChanged = False
    End If
    hDestDC = .hDC
    .ScaleMode = vbPixels
    .ForeColor = MENUFORCOLOR
    lWidth = CLng(.ScaleWidth)
    lHeight = CLng(.ScaleHeight)
End With
    
hSrcDC = picCache.hDC
    
If lMax = 0 Or hDestDC = 0 Or hSrcDC = 0 Then
    GoTo ExitPoint
End If
    
' first, paint the menus up to the currently select one
For l = 1 To mlMenuCur
    With picMenu
        ' draw the button
        BitBlt hDestDC, 0, (l - 1) * mlButtonHeight, _
                lWidth, mlButtonHeight, _
                hSrcDC, 0, 0, SRCCOPY
        ' print the caption
        sCaption = colMenus.Item(l).Caption
        .CurrentX = (lWidth \ 2) - (.TextWidth(sCaption) \ 2)
        .CurrentY = (l - 1) * mlButtonHeight + 2
        picMenu.Print sCaption
    End With
Next
    
' now, paint the menus below the currently seleted one (from the bottom up)
For l = lMax To mlMenuCur + 1 Step -1
    With picMenu
        ' draw the button
        BitBlt hDestDC, 0, _
                lHeight - (lMax - l + 1) * mlButtonHeight, _
                lWidth, mlButtonHeight, _
                hSrcDC, 0, 0, SRCCOPY
        ' print the caption
        sCaption = colMenus.Item(l).Caption
        .CurrentX = (lWidth \ 2) - (.TextWidth(sCaption) \ 2)
        .CurrentY = lHeight - (lMax - l + 1) * mlButtonHeight + 2
        picMenu.Print sCaption
    End With
Next
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
    ShowError Err.Number, Err.Description, "Repaint", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint
    
End Sub

' the new current menu is further down on the menu than the previous one
' we need to move the menus up from the previous menu + 1 to the new current menu
' support subroutine for Paint
Private Sub ReselectDown()
    Dim lStartY As Long
    Dim lStopY As Long
    Dim lTopOfGroupY As Long
    Dim lPixelCount As Long
    Dim lMax As Long
    Dim hDestDC As Long
    Dim hSrcDC As Long
    Dim lWidth As Long
    Dim bFirst As Boolean
    
    ' setup variables
    bFirst = True
    lMax = colMenus.Count
    With picMenu
        hDestDC = .hDC
        .ForeColor = vbButtonText
        lWidth = .ScaleWidth
        lStopY = mlMenuPrev * mlButtonHeight
        lStartY = .ScaleHeight - (lMax - mlMenuCur) * mlButtonHeight
        lTopOfGroupY = .ScaleHeight - (lMax - mlMenuPrev) * mlButtonHeight
    End With
    hSrcDC = picCache.hDC
    
    If lMax = 0 Or hDestDC = 0 Or hSrcDC = 0 Then
        Exit Sub
    End If
    Do
        BitBlt hDestDC, 0, lStopY, lWidth, _
            lStartY - lStopY - lPixelCount - PIXELS_PER_BITBLT - ((Not (bFirst)) * PIXELS_PER_BITBLT), _
            hDestDC, 0, lStopY + PIXELS_PER_BITBLT, SRCCOPY
        If bFirst Then
            BitBlt hDestDC, 0, lStartY - PIXELS_PER_BITBLT, _
                lWidth, PIXELS_PER_BITBLT, _
                hSrcDC, 0, mlButtonHeight + 3, SRCCOPY
            bFirst = False
        End If
        lPixelCount = lPixelCount + PIXELS_PER_BITBLT
    Loop Until lTopOfGroupY - ((lPixelCount + 1) * PIXELS_PER_BITBLT) <= lStopY

    ' make sure the group is in it's correct final position
    BitBlt hDestDC, 0, lStopY, lWidth, lStartY - lStopY - 1 - lPixelCount - (Not (bFirst) * PIXELS_PER_BITBLT), _
        hDestDC, 0, lTopOfGroupY - lPixelCount, SRCCOPY
        
End Sub

' the new current menu is further up on the menu than the previous one
' we need to move the menus down from the current menu + 1 to the previous menu
' support subroutine for Paint
Private Sub ReselectUp()
    Dim lStartY As Long
    Dim lStopY As Long
    Dim lBottomOfGroupY As Long
    Dim lPixelCount As Long
    Dim lMax As Long
    Dim hDestDC As Long
    Dim hSrcDC As Long
    Dim lWidth As Long
    Dim bFirst As Boolean
    
    ' setup variables
    bFirst = True
    lMax = colMenus.Count
    With picMenu
        hDestDC = .hDC
        .ForeColor = vbButtonText
        lWidth = .ScaleWidth
        lStartY = (mlMenuCur) * mlButtonHeight
        lStopY = .ScaleHeight - (lMax - mlMenuPrev) * mlButtonHeight
        lBottomOfGroupY = mlMenuPrev * mlButtonHeight
    End With
    hSrcDC = picCache.hDC
    
    If lMax = 0 Or hDestDC = 0 Or hSrcDC = 0 Then
        Exit Sub
    End If
    Do
        BitBlt hDestDC, 0, _
            lStartY + lPixelCount + PIXELS_PER_BITBLT + ((Not (bFirst)) * PIXELS_PER_BITBLT), _
            lWidth, _
            lStopY - lStartY - lPixelCount - PIXELS_PER_BITBLT - ((Not (bFirst)) * PIXELS_PER_BITBLT), _
            hDestDC, 0, lStartY + lPixelCount + ((Not (bFirst)) * PIXELS_PER_BITBLT), SRCCOPY
        If bFirst Then
            BitBlt hDestDC, 0, _
                lStartY + (lPixelCount * PIXELS_PER_BITBLT), _
                lWidth, PIXELS_PER_BITBLT, _
                hSrcDC, 0, mlButtonHeight + 3, SRCCOPY
            bFirst = False
        End If
        lPixelCount = lPixelCount + PIXELS_PER_BITBLT
    Loop Until lBottomOfGroupY + lPixelCount >= lStopY
    
    ' make sure the group is in it's correct final position
    BitBlt hDestDC, 0, _
        lStopY - (mlMenuPrev - mlMenuCur) * mlButtonHeight - PIXELS_PER_BITBLT, _
        lWidth, (mlMenuPrev - mlMenuCur) * mlButtonHeight + PIXELS_PER_BITBLT, _
        hDestDC, 0, lStartY + lPixelCount - PIXELS_PER_BITBLT, SRCCOPY
        
End Sub

' draw the icons for the currently select menu
' support subroutine for Paint
Private Sub DrawIcons()
    colMenus.Item(mlMenuCur).PaintItems IconStart(), mlMenuCur, ClipY(), colMenus.Count
End Sub

' support subroutine for Paint
Private Sub SetMenuButtonsHotSpot()
    Dim lIndex As Long
    Dim lMax As Long
    Dim VMenu As VMenu
    
    lMax = colMenus.Count
    For Each VMenu In colMenus
        With VMenu
            lIndex = .Index
            If lIndex <= mlMenuCur Then      ' the menu is at the top of the control
                .ButtonTop = (lIndex - 1) * mlButtonHeight
            Else                            ' the menu is at the bottom of the control
                .ButtonTop = picMenu.ScaleHeight - (lMax - lIndex + 1) * mlButtonHeight
            End If
        End With
    Next
End Sub

' determines if the mouse was clicked in a menu button
' returns the index of the menu clicked
' if no menu clicked, returns 0
Public Function IsMenuButtonClicked(ByVal ptX As Long, ByVal ptY As Long) As Long
    Dim VMenu As VMenu
    
    For Each VMenu In colMenus
        With VMenu
            If .IsMenuSelected(ptX, ptY) Then
                IsMenuButtonClicked = .Index
                Exit Function
            End If
        End With
    Next
End Function

Public Property Let NumberOfMenusChanged(ByVal bNewValue As Boolean)
    mbNumberOfMenusChanged = bNewValue
End Property

Public Property Get TotalMenuItems() As Long
    Dim VMenu As VMenu
    Dim lTotal As Long
    
    For Each VMenu In colMenus
        lTotal = lTotal + VMenu.MenuItemCount
    Next
    TotalMenuItems = lTotal
End Property

' Process mouse events
' Note that even if we get a hit, we must process all code
' The individual objects handle their own paints depending what
' the mouse is doing.
Public Function MouseProcess(ByVal iMousePosition As Integer, ByVal x As Long, ByVal y As Long, Optional lHitType As Long) As Long
    Dim lResult As Long
    Dim MenuItems As MenuItems
    Const HIT_TYPE_MENU_BUTTON = 1
    Const HIT_TYPE_MENUITEM = 2
    Const HIT_TYPE_UP_ARROW = 3
    Const HIT_TYPE_DOWN_ARROW = 4
    Const SCROLL_UP = 100
    Const SCROLL_DOWN = -100
    
    ' first test for a MenuButtonHit
    If iMousePosition = MOUSE_DOWN Then
        lResult = IsMenuButtonClicked(x, y)
        'if lResult is non-zero we have a hit
        ' so tell the calling procedure and return
        If lResult <> 0 Then
            lHitType = HIT_TYPE_MENU_BUTTON
            MouseProcess = lResult
        End If
    End If
    
    ' test for a MenuItemHit
    ' don't do the test if we got a MenuButtonHit
    If lResult = 0 Then
        Set MenuItems = colMenus.Item(mlMenuCur).MenuItems()
        lResult = MenuItems.MouseProcess(iMousePosition, x, y)
        If iMousePosition = MOUSE_UP And lResult > 0 Then
            lHitType = HIT_TYPE_MENUITEM
            MouseProcess = lResult
        End If
    End If
    ' test for arrow buttons
    lResult = colMenus.Item(mlMenuCur).MouseProcessForArrows(iMousePosition, x, y)
    If lResult <> 0 Then
        Select Case lResult
            Case SCROLL_UP
                ScrollUp
                lHitType = HIT_TYPE_UP_ARROW
            Case SCROLL_DOWN
                ScrollDown
                lHitType = HIT_TYPE_DOWN_ARROW
        End Select
    End If
End Function

Private Sub ScrollUp()
'******************************************************************************
'** SubRoutine.....: ScrollUp
'**
'** Description....: Move Icons and Caption Up
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 17-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** None
'******************************************************************************
On Error GoTo ErrorSection

Dim lStartY As Long
Dim lStopY As Long
Dim lTopOfGroupY As Long
Dim lPixelCount As Long
Dim lMax As Long
Dim hDestDC As Long
Dim hSrcDC As Long
Dim lWidth As Long
Dim lPixel As Long
    
With colMenus.Item(mlMenuCur)
    .TopMenuItem = .TopMenuItem - 1
    ' the upbutton is visible, hide it so it doesn't scroll
    .HideButton TYPE_UP, mlMenuCur
End With
    
' setup variables
lMax = colMenus.Count
With picMenu
    hDestDC = .hDC
    .ForeColor = vbButtonText
    lWidth = .ScaleWidth
    lStartY = mlMenuCur * mlButtonHeight
    lStopY = .ScaleHeight - (lMax - mlMenuCur) * mlButtonHeight
End With
hSrcDC = picCache.hDC
    
If lMax = 0 Or hDestDC = 0 Or hSrcDC = 0 Then
    GoTo ExitPoint
End If
    
For lPixel = 1 To SIZE_ICON * 2
    ' move the screen up
    BitBlt hDestDC, 0, _
            lStartY + PIXELS_PER_BITBLT_S, _
            lWidth, lStopY - lStartY - 2, _
            hDestDC, 0, lStartY, SRCCOPY
    ' repaint the background
    BitBlt hDestDC, 0, _
            lStartY, lWidth, _
            PIXELS_PER_BITBLT_S, _
            hSrcDC, 0, mlButtonHeight + 3, SRCCOPY
    lPixelCount = lPixelCount + PIXELS_PER_BITBLT_S
Next
    
DrawIcons
SetMenuButtonsHotSpot
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
    ShowError Err.Number, Err.Description, "ScrollUp", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Private Sub ScrollDown()
'******************************************************************************
'** SubRoutine.....: ScrollDown
'**
'** Description....: Move Icons and Caption Down
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 17-Jul-2001.
'**
'** Modifications..:
'**
'** Arguments
'** Name                Type     Acces   Description
'** ------------------  -------  ------  -------------------------------------
'** None
'******************************************************************************
On Error GoTo ErrorSection

Dim lStartY As Long
Dim lStopY As Long
Dim lTopOfGroupY As Long
Dim lPixelCount As Long
Dim lMax As Long
Dim hDestDC As Long
Dim hSrcDC As Long
Dim lWidth As Long
Dim lPixel As Long

With colMenus.Item(mlMenuCur)
    .TopMenuItem = .TopMenuItem + 1
    ' the down button is visible, hide it so it doesn't scroll
    .HideButton TYPE_DOWN, colMenus.Count - mlMenuCur
End With
    
'** setup variables
lMax = colMenus.Count
With picMenu
    hDestDC = .hDC
    .ForeColor = vbButtonText
    lWidth = .ScaleWidth
    lStopY = mlMenuCur * mlButtonHeight
    lStartY = .ScaleHeight - (lMax - mlMenuCur) * mlButtonHeight
End With
hSrcDC = picCache.hDC
    
If lMax = 0 Or hDestDC = 0 Or hSrcDC = 0 Then
        GoTo ExitPoint
End If
        
For lPixel = 1 To SIZE_ICON * 2
    '** move the screen up
    BitBlt hDestDC, 0, lStopY, _
            lWidth, lStartY - lStopY, _
            hDestDC, 0, lStopY + PIXELS_PER_BITBLT_S, SRCCOPY
    '** repaint the background
    BitBlt hDestDC, 0, _
            lStartY - 1, lWidth, PIXELS_PER_BITBLT_S, _
            hSrcDC, 0, mlButtonHeight + 3, SRCCOPY
    lPixelCount = lPixelCount + PIXELS_PER_BITBLT_S
Next
    
'** redraw the icons
DrawIcons
SetMenuButtonsHotSpot
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
    ShowError Err.Number, Err.Description, "ScrollDown", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint

End Sub

Private Function IconStart() As Long
    Dim l As Long
    Dim lIconStart As Long
    
    ' calculate the offset for our first icon
    For l = 1 To mlMenuCur - 1
        lIconStart = lIconStart + colMenus.Item(l).MenuItemCount
    Next
    IconStart = lIconStart + colMenus.Item(mlMenuCur).TopMenuItem - 1
End Function

Private Function ClipY() As Long
    
    ' calculate the clipping area where the menu bottoms start at the bottom of picmenu
    With picMenu
        ClipY = .ScaleHeight - ((colMenus.Count) - mlMenuCur) * mlButtonHeight
    End With
End Function



