VERSION 5.00
Begin VB.UserControl CtlVerticalMenu 
   Alignable       =   -1  'True
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "VerticalMenu.ctx":0000
   PropertyPages   =   "VerticalMenu.ctx":030A
   ScaleHeight     =   57
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   ToolboxBitmap   =   "VerticalMenu.ctx":0347
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   600
      ScaleHeight     =   420
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picCache 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1320
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   1920
      Picture         =   "VerticalMenu.ctx":0659
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   1920
      Picture         =   "VerticalMenu.ctx":0B9B
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "CtlVerticalMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'******************************************************************************
'** Class Name.....: CtlVerticalMenu
'** Description....: Acting like Outlook Menu
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
Private Const m_ClassName = "CtlVerticalMenu"

Private mMenus   As Menus
Private Const DI_NORMAL = &H3

Enum Icon_Size
    [16x16] = 0
    [32x32] = 1
    [48x48] = 2
    [64x64] = 3
End Enum

'Default Property Values
Private Const mclMenusMax           As Long = 1
Private Const mclMenuCur            As Long = 1
Private Const mclMenuStartup        As Long = 1
Private Const mcsMenuCaption        As String = "Menu"
Private Const mcsMenuItemCaption    As String = "Item"
Private Const mclMenuItemsMax       As Long = 1
Private Const mclMenuItemCur        As Long = 1

'Property Variables:
Dim m_MenuItemForeColor As OLE_COLOR
Dim m_MenuForeColor As OLE_COLOR
Dim m_BackEffect As Integer
Dim m_IconSize As Integer

Private mlMenusMax                  As Long
Private mlMenuCur                   As Long
Private mlMenuStartup               As Long
Private msMenuCaption               As String
Private msMenuItemCaption           As String
Private mpicMenuItemIcon            As Picture
Private mlMenuItemsMax              As Long
Private mlMenuItemCur               As Long
Private mbInitializing              As Boolean

' Constants
Private Const HIT_TYPE_MENU_BUTTON  As Integer = 1
Private Const HIT_TYPE_MENUITEM     As Integer = 2
Private Const HIT_TYPE_UP_ARROW     As Integer = 3
Private Const HIT_TYPE_DOWN_ARROW   As Integer = 4
Private Const BUTTON_HEIGHT         As Integer = 18
Private Const MOUSE_IN_CAPTION      As Integer = -2

'Event Declarations:
Public Event Show()
Public Event Resize()
Public Event Hide()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Paint()
Public Event MenuItemClick(MenuNumber As Long, MenuItem As Long)
Public Event MenuClick(ByVal MenuNumber As Long)

'** API
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Const m_def_IconSize = 1
Const m_def_BackEffect = 0
'Default Property Values:
Const m_def_MenuItemForeColor = vbWhite



'**************************
'****    Properties    ****
'**************************

Public Property Get MenuItemForeColor() As OLE_COLOR
    MenuItemForeColor = m_MenuItemForeColor
End Property

Public Property Let MenuItemForeColor(ByVal New_MenuItemForeColor As OLE_COLOR)
    m_MenuItemForeColor = New_MenuItemForeColor
    PropertyChanged "MenuItemForeColor"
    ITEMFORCOLOR = m_MenuItemForeColor
    picMenu.Cls
    UserControl_Paint
End Property

Public Property Get MenuForeColor() As OLE_COLOR
    MenuForeColor = m_MenuForeColor
End Property

Public Property Let MenuForeColor(ByVal New_MenuForeColor As OLE_COLOR)
    m_MenuForeColor = New_MenuForeColor
    PropertyChanged "MenuForeColor"
    MENUFORCOLOR = m_MenuForeColor
    picMenu.Cls
    UserControl_Paint
End Property

Public Property Get MenusMax() As Long
  MenusMax = mlMenusMax
End Property

Public Property Let MenusMax(ByVal alMenusMax As Long)
  Dim lIndex      As Long
  Dim lSavMenuCur As Long
  
  'Check for maximum menus allowed (8)
  If alMenusMax <= 0 Then
    alMenusMax = 1
  End If
  If alMenusMax > 8 Then
    alMenusMax = 8
  End If
  'Paint menus
  Select Case alMenusMax
    Case mlMenusMax             'Nothing to do
    Case Is > mlMenusMax        'Add menus
      lSavMenuCur = mlMenuCur
      For mlMenuCur = mlMenusMax + 1 To alMenusMax
        With mMenus
          'Add menu and set caption
          .Add "", mlMenuCur, picMenu
          MenuCaption = mcsMenuCaption & CStr(mlMenuCur)
          'Set the up/down bitmaps
          Set .Item(mlMenuCur).UpBitmap = imgUp.Picture
          Set .Item(mlMenuCur).DownBitmap = imgDown.Picture
          Set .Item(mlMenuCur).ImageCache = picCache
          'Add MenuItems to the menu
          .Item(mlMenuCur).AddMenuItem mcsMenuItemCaption, 1, mpicMenuItemIcon
        End With
      Next mlMenuCur
      mlMenuCur = lSavMenuCur
    Case Is < mlMenusMax        'Delete menus
      For lIndex = mlMenusMax To alMenusMax + 1 Step -1
        With mMenus
          .Delete lIndex
          If alMenusMax < mlMenuCur Then MenuCur = alMenusMax
        End With
      Next lIndex
  End Select
  'Save new state of control
  mlMenusMax = alMenusMax
  mMenus.NumberOfMenusChanged = True
  SetupCache
  UserControl_Paint
  'Save the Property
  PropertyChanged "MenusMax"
End Property

Public Property Get IconSize() As Icon_Size
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_IconSize As Icon_Size)
    m_IconSize = New_IconSize
    PropertyChanged "IconSize"
    UserControl.Cls
    picMenu.Cls
    FitIcon
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = picMenu.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picMenu.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    BACKGROUND_COLOR = picMenu.BackColor
    picCache.BackColor = BACKGROUND_COLOR
    SetupCache
    UserControl_Paint
End Property

Public Property Get MenuCur() As Long
  MenuCur = mlMenuCur
End Property

Public Property Let MenuCur(ByVal alMenuCur As Long)
'******************************************************************************
'** Description....: Set the current menu
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  mlMenuCur = alMenuCur
  mlMenuItemCur = 1           'Reset the menuitem
  With mMenus
    .MenuCur = mlMenuCur
    mlMenuItemsMax = .Item(mlMenuCur).MenuItemCount
    MenuCaption = .Item(mlMenuCur).Caption
  End With
  'Save the Property
  PropertyChanged "MenuCur"
End Property

Public Property Get MenuStartup() As Long
'******************************************************************************
'** Description....: Defines the menu to show at startup
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  MenuStartup = mlMenuStartup
End Property

Public Property Let MenuStartup(ByVal alMenuStartup As Long)
  mlMenuStartup = alMenuStartup
  PropertyChanged "MenuStartup"
End Property


Public Property Get MenuCaption() As String
'******************************************************************************
'** Description....: Get the value
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  On Error Resume Next
  MenuCaption = msMenuCaption
End Property

Public Property Let MenuCaption(ByVal asMenuCaption As String)
'******************************************************************************
'** Description....: Defines the Caption of the current menu
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  msMenuCaption = asMenuCaption
  mMenus.Item(mlMenuCur).Caption = asMenuCaption
  'Force paint
  UserControl_Paint
  'Save the Property
  PropertyChanged "MenuCaption"
End Property

Public Property Get MenuItemCaption() As String
  msMenuItemCaption = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Caption
  MenuItemCaption = msMenuItemCaption
End Property

Public Property Let MenuItemCaption(ByVal asMenuItemCaption As String)
'******************************************************************************
'** Description....: Defines the Caption of the current menuitem
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  With mMenus.Item(mlMenuCur)
    .MenuItemItem(mlMenuItemCur).Caption = asMenuItemCaption
    msMenuItemCaption = asMenuItemCaption
  End With
  'Repaint the control
  If Not mbInitializing Then
    picMenu.Cls
    UserControl_Paint
  End If
  PropertyChanged "MenuItemCaption"
End Property

Public Property Get MenuItemIcon() As Picture
  Set MenuItemIcon = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Button
End Property

Public Property Set MenuItemIcon(ByVal New_MenuItemIcon As Picture)
'******************************************************************************
'** Description....: Defines the icon of the current menuitem
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  Set mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Button = New_MenuItemIcon
  'Repaint the control
  If Not mbInitializing Then
    SetupCache
    UserControl_Paint
  End If
  PropertyChanged "MenuItemIcon"
End Property

Public Property Get MenuItemKey() As String
  MenuItemKey = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Key
End Property

Public Property Let MenuItemKey(ByVal New_MenuItemKey As String)
'******************************************************************************
'** Description....: Defines the key of the current menuitem
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Key = New_MenuItemKey
  PropertyChanged "MenuItemKey"
End Property

Public Property Get MenuItemTag() As String
  MenuItemTag = mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Tag
End Property

Public Property Let MenuItemTag(ByVal New_MenuItemTag As String)
'******************************************************************************
'** Description....: Defines the tag of the current menuitem
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  mMenus.Item(mlMenuCur).MenuItemItem(mlMenuItemCur).Tag = New_MenuItemTag
  PropertyChanged "MenuItemTag"
End Property

Public Property Get MenuItemsMax() As Long
  MenuItemsMax = mlMenuItemsMax
End Property

Public Property Let MenuItemsMax(ByVal alMenuItemsMax As Long)
'******************************************************************************
'** Description....: maximum of menuitems for the current menu
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  
  Dim lSavMenuItemCur As Long
  'Test for maximum of menuentries
  If alMenuItemsMax < 0 Or alMenuItemsMax > 12 Then
    Beep
    MsgBox "MenuItemsMax must be between 0 and 12", vbOKOnly
    Exit Property
  End If
  
  'Build the menu
  lSavMenuItemCur = mlMenuItemCur
  Select Case alMenuItemsMax
    Case mlMenuItemsMax             'Nothing to do
    Case Is > mlMenuItemsMax        'Add menus
      With mMenus.Item(mlMenuCur)
        For mlMenuItemCur = mlMenuItemsMax + 1 To alMenuItemsMax
          .AddMenuItem mcsMenuItemCaption, mlMenuItemCur, mpicMenuItemIcon
          MenuItemCaption = mcsMenuItemCaption & CStr(mlMenuItemCur)
        Next mlMenuItemCur
        mlMenuItemCur = lSavMenuItemCur
      End With
    Case Is < mlMenuItemsMax        'Delete menus
      With mMenus.Item(mlMenuCur)
        For mlMenuItemCur = mlMenuItemsMax To alMenuItemsMax + 1 Step -1
          .DeleteMenuItem mlMenuItemCur
        Next mlMenuItemCur
        mlMenuItemCur = lSavMenuItemCur
        If alMenuItemsMax < mlMenuItemCur Then mlMenuItemCur = alMenuItemsMax
      End With
  End Select
  'Reset the caption in the properties window
  mlMenuItemsMax = alMenuItemsMax
  'Repaint the control
  SetupCache
  picMenu.Refresh
  UserControl_Paint
  PropertyChanged "MenuItemsMax"
End Property

Public Property Get MenuItemCur() As Long
  MenuItemCur = mlMenuItemCur
End Property

Public Property Let MenuItemCur(ByVal alMenuItemCur As Long)
'******************************************************************************
'** Description....: Defines the current menuitem
'** Author, date...: Yves Lessard  17-Jul-2001.
'******************************************************************************
  'Test for correctness
  If alMenuItemCur > mlMenuItemsMax Then
    Beep
    MsgBox "The current item must be between 0 and MenuItemsMax", vbOKOnly
    Exit Property
  End If
  mlMenuItemCur = alMenuItemCur
  PropertyChanged "MenuItemCur"
End Property


'********************************
'****    Private Methodes    ****
'********************************



Private Sub FitIcon()
    Select Case m_IconSize
        Case Icon_Size.[16x16]
            SIZE_ICON = 16
        Case Icon_Size.[32x32]
            SIZE_ICON = 32
        Case Icon_Size.[48x48]
            SIZE_ICON = 48
        Case Icon_Size.[64x64]
            SIZE_ICON = 64
    End Select
    SetupCache
    UserControl_Paint
End Sub

Private Sub picCache_Resize()
  DrawCacheMenuButton
End Sub

Private Sub picMenu_DblClick()
  Dim recPOINTAPI As POINTAPI
  Dim lResCod     As Long
  'If picMenu considers a second mousedown event as a dblclick, the
  'MouseDown event does not file so we need to do it here instead
  lResCod = GetCursorPos(recPOINTAPI)
  lResCod = ScreenToClient(picMenu.hWnd, recPOINTAPI)
  picMenu_MouseDown vbLeftButton, 0, CSng(recPOINTAPI.x), CSng(recPOINTAPI.y)
End Sub

Private Sub picMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lIndex    As Long
  Dim lHitType  As Long
  If Button = vbLeftButton Then
    With mMenus
      'Care only about MenuButton hits; all others are already processed
      lIndex = .MouseProcess(MOUSE_DOWN, CLng(x), CLng(y), lHitType)
      If lHitType = HIT_TYPE_MENU_BUTTON And lIndex > 0 Then
        If MenuCur <> lIndex Then
            MenuCur = lIndex
            RaiseEvent MenuClick(MenuCur)
        End If
      End If
    End With
  End If
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  mMenus.MouseProcess MOUSE_MOVE, CLng(x), CLng(y)
End Sub

Private Sub picMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim lMenuItem As Long
  Dim lHitType  As Long
  If Button = vbLeftButton Then
    lMenuItem = mMenus.MouseProcess(MOUSE_UP, CLng(x), CLng(y), lHitType)
    If lHitType = HIT_TYPE_MENUITEM And lMenuItem > 0 Then
      picMenu_MouseMove Button, Shift, x, y
      RaiseEvent MenuItemClick(mlMenuCur, lMenuItem)
      picMenu_MouseMove 0, 0, 0, 0
    End If
  End If
End Sub

Private Sub picMenu_Paint()
  If picMenu.Visible Then
    mMenus.Paint
  End If
End Sub

Private Sub UserControl_Paint()
  If Not mbInitializing Then picMenu_Paint
End Sub

Private Sub UserControl_Initialize()
  Set mMenus = New Menus
  Set mMenus.Menu = picMenu
  Set mMenus.Cache = picCache
End Sub

Private Sub UserControl_Resize()
  'scale menu
  With picMenu
    .Left = 0
    .Top = 0
    .Width = UserControl.ScaleWidth
    .Height = UserControl.ScaleHeight
  End With
  'scale workspace
  With picCache
    .Width = picMenu.Width
    .Height = (BUTTON_HEIGHT * 2) + 33
  End With
  'paint control
  picMenu.Refresh
End Sub

Private Sub UserControl_Terminate()
  'Destroy menus
  Set mMenus = Nothing
End Sub

Public Sub Refresh()
'Function:  Refreshs the control
'Arguments: -
'Versions:
'Author   Date        Remark
'--------------------------------------------------------------------------
'psch     14.12.1998  First build
  
  UserControl_Paint

End Sub


Public Sub SetupCache()
'******************************************************************************
'** SubRoutine.....: SetupCache
'**
'** Description....: Sets the cache (picturebox) up
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

Dim lMenuItemCount  As Long
Dim lMIndex         As Long
Dim lMIIndex        As Long
Dim lIconIndex      As Long
Dim I_OFFSET As Long
  
I_OFFSET = BUTTON_HEIGHT * 2 + SIZE_ICON
'Initialise the control
picCache.Cls
DrawCacheMenuButton
'Total MenuItems on the control
lMenuItemCount = mMenus.TotalMenuItems
With picCache
    'Set the height for a menu button, space for an unpainted button space for an
    'unpainted icon and all the MenuItem icons
    .Height = BUTTON_HEIGHT * 2 + (lMenuItemCount + 1) * SIZE_ICON
    'Get each icon for each menuitem
    lIconIndex = 0
    For lMIndex = 1 To mMenus.Count
      For lMIIndex = 1 To mMenus.Item(lMIndex).MenuItemCount
        lIconIndex = lIconIndex + 1
        DrawIconEx picCache.hDC, 0, I_OFFSET + (lIconIndex - 1) * SIZE_ICON, _
                    mMenus.Item(lMIndex).MenuItemItem(lMIIndex).Button.Handle, _
                    SIZE_ICON, SIZE_ICON, 0, 0, DI_NORMAL
      Next lMIIndex
    Next lMIndex
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
    ShowError Err.Number, Err.Description, "SetupCache", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint
  
End Sub

Private Sub ProcessDefaultIcon()
'******************************************************************************
'** SubRoutine.....: ProcessDefaultIcon
'**
'** Description....: Sets the default icon for the current menuitem
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 17-Jul-2001.
'**
'** Modifications..:
'**
'******************************************************************************
  If mpicMenuItemIcon Is Nothing Then Set mpicMenuItemIcon = UserControl.Picture
  UserControl.Picture = LoadPicture()
End Sub

Private Sub DrawCacheMenuButton()
'******************************************************************************
'** SubRoutine.....: DrawCacheMenuButton
'**
'** Description....: Draws the menubuttons
'**
'** Cie/Co ....: SevySoft
'** Author, date...: Yves Lessard , 17-Jul-2001.
'**
'** Modifications..:
'**
'******************************************************************************
On Error GoTo ErrorSection

Dim recRECT As RECT
'Set properties
With recRECT
    .Left = 0
    .Top = 0
    .Right = picCache.ScaleWidth
    .Bottom = BUTTON_HEIGHT
End With
'Call API-Drawing-Routine
DrawEdge picCache.hDC, recRECT, BDR_RAISED, BF_RECT Or BF_MIDDLE
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
    ShowError Err.Number, Err.Description, "DrawCacheMenuButton", m_ClassName, vbLogEventTypeError
End Select
Resume ExitPoint
End Sub

Private Sub UserControl_InitProperties()
  
  'We're initializing now
  mbInitializing = True
  'Set button height for icons
  mMenus.ButtonHeight = BUTTON_HEIGHT
  m_MenuForeColor = Ambient.ForeColor
  'Set property defaults
  m_IconSize = m_def_IconSize
  SIZE_ICON = 32
  'Set the default icon
  ProcessDefaultIcon
  'Setup the image cache
  BACKGROUND_COLOR = Ambient.BackColor
  With picCache
    .Width = picMenu.Width
    .Height = (BUTTON_HEIGHT * 2) + 33
    .BackColor = BACKGROUND_COLOR
  End With
  picMenu.BackColor = BACKGROUND_COLOR
  
  'Setup the control
  MenusMax = mclMenusMax
  MenuCur = mclMenuStartup
  MenuStartup = mclMenuStartup
  'Setup the menu caption button and menu item icon cache
  SetupCache
  'We're finished initializing
  mbInitializing = False
  m_BackEffect = m_def_BackEffect
    m_MenuForeColor = Ambient.ForeColor
    m_MenuItemForeColor = m_def_MenuItemForeColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Dim lSavMenuItemCur As Long
  'We're initializing now
  mbInitializing = True
  'Read the properies out of their storage
  m_MenuForeColor = PropBag.ReadProperty("MenuForeColor", Ambient.ForeColor)
  MENUFORCOLOR = m_MenuForeColor
  m_MenuItemForeColor = PropBag.ReadProperty("MenuItemForeColor", m_def_MenuItemForeColor)
  ITEMFORCOLOR = m_MenuItemForeColor
  With PropBag
    'mbEnabled = .ReadProperty("Enabled", mclEnabled)
    mlMenuItemCur = mclMenuItemCur
    mlMenuItemsMax = mclMenuItemsMax
    'Set the icon
    m_IconSize = .ReadProperty("IconSize", m_def_IconSize)
    FitIcon
    Set mpicMenuItemIcon = .ReadProperty("MenuItemIcon0", Nothing)
    ProcessDefaultIcon
    'Setup the image cache
    With picCache
      .Width = UserControl.Width
      .Height = (BUTTON_HEIGHT * 2) + 33
    End With

    'Add the first menu (which already exists on the form) to the collection.
    '(Note that calling MenusMax only adds and deletes menus other that the
    ' first item in the collection)
    mMenus.ButtonHeight = BUTTON_HEIGHT
    MenusMax = .ReadProperty("MenusMax", mclMenusMax)
    'Setup the control arrays
    For mlMenuCur = 1 To mlMenusMax
      'Current menu
      MenuCur = mlMenuCur
      msMenuCaption = .ReadProperty("MenuCaption" & CStr(mlMenuCur), mcsMenuCaption)
      MenuCaption = msMenuCaption
      'Maxmimum menus
      MenuItemsMax = .ReadProperty("MenuItemsMax" & CStr(mlMenuCur), mclMenuItemsMax)
      'Current menuitem
      lSavMenuItemCur = mlMenuItemCur
      For mlMenuItemCur = 1 To mMenus.Item(mlMenuCur).MenuItemCount
        Set MenuItemIcon = .ReadProperty("MenuItemIcon" & CStr(mlMenuCur) & CStr(mlMenuItemCur), mpicMenuItemIcon)
        MenuItemCaption = .ReadProperty("MenuItemCaption" & CStr(mlMenuCur) & CStr(mlMenuItemCur), mcsMenuItemCaption)
        MenuItemKey = .ReadProperty("MenuItemKey" & CStr(mlMenuCur) & CStr(mlMenuItemCur), "")
        MenuItemTag = .ReadProperty("MenuItemTag" & CStr(mlMenuCur) & CStr(mlMenuItemCur), "")
      Next mlMenuItemCur
      mlMenuItemCur = lSavMenuItemCur
    Next mlMenuCur
    
    'Reset mlMenuCur right away so we don't have errors
    mlMenuCur = .ReadProperty("MenuCur", mclMenuCur)
    'Read all other properies
    MenuItemCur = mclMenuItemCur
    mlMenuStartup = .ReadProperty("MenuStartup", mclMenuStartup)
    MenuStartup = mlMenuStartup
    MenuCur = mlMenuStartup
  End With
  picMenu.BackColor = PropBag.ReadProperty("BackColor", &H8000000C)
  BACKGROUND_COLOR = picMenu.BackColor
  picCache.BackColor = BACKGROUND_COLOR
  'Setup the menu caption button and menu item icon cache
  SetupCache
  'We're finished initializing
  mbInitializing = False
  
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
  Dim lSavMenuCur     As Long
  Dim lSavMenuItemCur As Long
  
  With PropBag
    'Save general properties
    Call .WriteProperty("MenusMax", mlMenusMax, mclMenusMax)
    Call .WriteProperty("MenuCur", mlMenuCur, mclMenuCur)
    Call .WriteProperty("MenuStartup", mlMenuStartup, mclMenuStartup)
    'Save menu properites
    lSavMenuCur = mlMenuCur
    For mlMenuCur = 1 To mlMenusMax
      Call .WriteProperty("MenuCaption" & CStr(mlMenuCur), mMenus.Item(mlMenuCur).Caption, mcsMenuCaption)
      'Save menuitem properties
      Call .WriteProperty("MenuItemsMax" & CStr(mlMenuCur), mMenus.Item(mlMenuCur).MenuItemCount, mclMenuItemsMax)
      lSavMenuItemCur = mlMenuItemCur
      For mlMenuItemCur = 1 To mMenus.Item(mlMenuCur).MenuItemCount
        Call .WriteProperty("MenuItemIcon" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemIcon, Nothing)
        Call .WriteProperty("MenuItemCaption" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemCaption, mcsMenuItemCaption)
        Call .WriteProperty("MenuItemKey" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemKey, "")
        Call .WriteProperty("MenuItemTag" & CStr(mlMenuCur) & CStr(mlMenuItemCur), MenuItemTag, "")
      Next mlMenuItemCur
      mlMenuItemCur = lSavMenuItemCur
    Next mlMenuCur
    mlMenuCur = lSavMenuCur
    'Save the other properties
    Call .WriteProperty("MenuItemIcon0", mpicMenuItemIcon, mpicMenuItemIcon)
  End With
    Call PropBag.WriteProperty("BackColor", picMenu.BackColor, &H8000000C)
    Call PropBag.WriteProperty("IconSize", m_IconSize, m_def_IconSize)
    Call PropBag.WriteProperty("MenuForeColor", m_MenuForeColor, Ambient.ForeColor)
    Call PropBag.WriteProperty("MenuItemForeColor", m_MenuItemForeColor, m_def_MenuItemForeColor)
End Sub





