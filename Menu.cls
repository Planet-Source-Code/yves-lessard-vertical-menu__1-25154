VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MenuItems"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private msCaption As String      ' caption of the Menu
Private mlIndex As Long          ' location of the Menu
Private picMenu As PictureBox
Private picCache As PictureBox
Private mlButtonHeight As Long
Private mMenuItems As MenuItems
Private mpicUp As clsArrow
Private mpicDown As clsArrow
Private mHotSpot As RECT
Private mlTopMenuItemDisplayed As Long

Const TYPE_UP = 1
Const TYPE_DOWN = -1
Const BTN_UP = 1
Const BTN_DOWN = -1
Const SCROLL_DOWN = -100
Const SCROLL_UP = 100

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

Public Property Get Control() As Object
    Set Control = picMenu
End Property

Public Property Set Control(pic As Object)
    Set picMenu = pic
    
    ' also tell the arrows who the parent is
    Set mpicUp.ctlParent = pic
    Set mpicDown.ctlParent = pic
End Property

Public Function AddMenuItem(sCaption As String, lMenuItemlIndex As Long, picIcon As Object) As MenuItems
    With mMenuItems
        .Add sCaption, lMenuItemlIndex, mlButtonHeight, picIcon
        Set .Item(lMenuItemlIndex).Parent = picMenu
        Set .Item(lMenuItemlIndex).Cache = picCache
    End With
End Function

Public Sub DeleteMenuItem(lMenuItemlIndex As Long)
    mMenuItems.Delete lMenuItemlIndex
End Sub

Public Function MenuItemCount() As Long
    MenuItemCount = mMenuItems.Count
End Function

Public Function MenuItemItem(lMenuItemlIndex As Long) As MenuItem
    Set MenuItemItem = mMenuItems.Item(lMenuItemlIndex)
End Function

' process mouse events for arrow buttons
Public Function MouseProcessForArrows(ByVal iMousePosition, ByVal x As Long, ByVal y As Long) As Long
    Dim bResult As Boolean
    Dim pic As clsArrow
    Dim i As Integer
    Static lLastPosition(1) As Long
        
    For i = 0 To 1
        If i = 0 Then
            Set pic = mpicDown
        Else
            Set pic = mpicUp
        End If

        bResult = pic.bHitTest(iMousePosition, x, y)
        If bResult Then
            Select Case iMousePosition
                Case MOUSE_UP
                    If lLastPosition(i) = BTN_DOWN Then
                        If i = 0 Then
                            MouseProcessForArrows = SCROLL_DOWN
                        Else
                            MouseProcessForArrows = SCROLL_UP
                        End If
                    End If
                    lLastPosition(i) = iMousePosition
                Case MOUSE_DOWN
                    lLastPosition(i) = iMousePosition
                Case MOUSE_MOVE
                    If lLastPosition(i) <> BTN_DOWN Then
                        lLastPosition(i) = iMousePosition
                    End If
            End Select
        Else
            If iMousePosition = MOUSE_UP Then
                lLastPosition(i) = BTN_UP
            End If
        End If
    Next
    Set pic = Nothing
End Function

Public Property Get ButtonHeight() As Long
    ButtonHeight = mlButtonHeight
End Property

Public Property Let ButtonHeight(ByVal lNewValue As Long)
    mlButtonHeight = lNewValue
    mpicUp.lButtonHeight = lNewValue
    mpicDown.lButtonHeight = lNewValue
End Property

Private Sub Class_Initialize()
    Set mMenuItems = New MenuItems
    
    ' create our up arrow
    Set mpicUp = New clsArrow
    mpicUp.iArrowType = TYPE_UP
    
    ' create our down arrow
    Set mpicDown = New clsArrow
    mpicDown.iArrowType = TYPE_DOWN
    
    mlTopMenuItemDisplayed = 1
End Sub

Private Sub Class_Terminate()
    Set mpicDown = Nothing
    Set mpicUp = Nothing
    Set picMenu = Nothing
End Sub

Public Property Get UpBitmap() As Object
    Set UpBitmap = mpicUp.objBitmap
End Property

Public Property Set UpBitmap(ByVal oNewValue As Object)
    Set mpicUp.objBitmap = oNewValue
End Property

Public Property Get DownBitmap() As Object
    Set DownBitmap = mpicDown.objBitmap
End Property

Public Property Set DownBitmap(ByVal oNewValue As Object)
    Set mpicDown.objBitmap = oNewValue
End Property

Public Property Set ImageCache(ByVal ctlNewValue As Object)
    Set picCache = ctlNewValue
End Property

' hittest to see if the points are in the menu button
Public Function IsMenuSelected(ByVal ptX As Long, ByVal ptY As Long) As Boolean
    IsMenuSelected = Not (PtInRect(mHotSpot, ptX, ptY) = 0)
    If Err.Number <> 0 Then
        IsMenuSelected = False
        Err.Clear
    End If
End Function

' menu button location
' all we need to do to set the structure is pass the top
' because we can compute the other locations
' same reason all we need to do is return the top location
Public Property Get ButtonTop() As Long
    ButtonTop = mHotSpot.Top
End Property

Public Property Let ButtonTop(ByVal lNewValue As Long)
    With picMenu
        .ScaleMode = vbPixels
        mHotSpot.Left = 0
        mHotSpot.Top = lNewValue
        mHotSpot.Right = .ScaleWidth
        mHotSpot.Bottom = lNewValue + mlButtonHeight
    End With
End Property

Public Function PaintItems(lIconStart As Long, lMenuCur As Long, lClipY As Long, lMax As Long) As Boolean
    Dim i As Integer
    
    If Not mMenuItems.Paint(mlTopMenuItemDisplayed, lIconStart, lMenuCur, lClipY) Then
        ' the second parameter for the down button is the
        ' number of buttons at the bottom of the menu
        mpicDown.Show alMenusAtBottom:=lMax - lMenuCur + 1
    Else
        mpicDown.Hide
    End If
    If mlTopMenuItemDisplayed > 1 Then
        ' the second parameter for the down button is the
        ' number of buttons at the to of the menu
        mpicUp.Show alMenusAtTop:=lMenuCur
    Else
        mpicUp.Hide
    End If
End Function

Public Property Get MenuItems() As MenuItems
    Set MenuItems = mMenuItems
End Property

Public Sub HideButton(iThisButton As Integer, lOffset As Long)
    If iThisButton = TYPE_UP Then
        mpicUp.Hide
    Else
        mpicDown.Hide
    End If
End Sub

Public Property Get TopMenuItem() As Long
    If mlTopMenuItemDisplayed = 0 Then
        mlTopMenuItemDisplayed = 1
    End If
    TopMenuItem = mlTopMenuItemDisplayed
End Property

Public Property Let TopMenuItem(ByVal lNewValue As Long)
    If lNewValue <> 0 Then
        mlTopMenuItemDisplayed = lNewValue
    End If
End Property

