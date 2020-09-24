VERSION 5.00
Object = "{68E32E84-1C15-483D-AA61-395296007689}#2.0#0"; "VerticalMenu.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VerticalMenu.CtlVerticalMenu CtlVerticalMenu1 
      Align           =   3  'Align Left
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   10769
      MenusMax        =   5
      MenuItemsMax1   =   4
      MenuItemIcon11  =   "Form1.frx":0000
      MenuItemCaption11=   "Browse CD "
      MenuItemIcon12  =   "Form1.frx":031A
      MenuItemCaption12=   "Send To Desktop"
      MenuItemIcon13  =   "Form1.frx":0634
      MenuItemCaption13=   "Config Computer"
      MenuItemIcon14  =   "Form1.frx":094E
      MenuItemCaption14=   "Play CD Music"
      MenuCaption2    =   "Menu2"
      MenuItemsMax2   =   6
      MenuItemIcon21  =   "Form1.frx":0C68
      MenuItemCaption21=   "Dial Up"
      MenuItemIcon22  =   "Form1.frx":0F82
      MenuItemCaption22=   "Log Off"
      MenuItemIcon23  =   "Form1.frx":129C
      MenuItemCaption23=   "Item3"
      MenuItemIcon24  =   "Form1.frx":15B6
      MenuItemCaption24=   "Item4"
      MenuItemIcon25  =   "Form1.frx":18D0
      MenuItemCaption25=   "Item5"
      MenuItemIcon26  =   "Form1.frx":1BEA
      MenuItemCaption26=   "Item6"
      MenuCaption3    =   "Menu3"
      MenuItemsMax3   =   12
      MenuItemIcon31  =   "Form1.frx":1F04
      MenuItemIcon32  =   "Form1.frx":221E
      MenuItemCaption32=   "Item2"
      MenuItemIcon33  =   "Form1.frx":2538
      MenuItemCaption33=   "Item3"
      MenuItemIcon34  =   "Form1.frx":2852
      MenuItemCaption34=   "Item4"
      MenuItemIcon35  =   "Form1.frx":2B6C
      MenuItemCaption35=   "Item5"
      MenuItemIcon36  =   "Form1.frx":2E86
      MenuItemCaption36=   "Item6"
      MenuItemIcon37  =   "Form1.frx":31A0
      MenuItemCaption37=   "Item7"
      MenuItemIcon38  =   "Form1.frx":34BA
      MenuItemCaption38=   "Item8"
      MenuItemIcon39  =   "Form1.frx":37D4
      MenuItemCaption39=   "Item9"
      MenuItemIcon310 =   "Form1.frx":3AEE
      MenuItemCaption310=   "Item10"
      MenuItemIcon311 =   "Form1.frx":3E08
      MenuItemCaption311=   "Item11"
      MenuItemIcon312 =   "Form1.frx":4122
      MenuItemCaption312=   "Item12"
      MenuCaption4    =   "Menu4"
      MenuItemIcon41  =   "Form1.frx":443C
      MenuCaption5    =   "Menu5"
      MenuItemIcon51  =   "Form1.frx":4756
      BackColor       =   11316313
      MenuForeColor   =   3092224
      MenuItemForeColor=   16384
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CtlVerticalMenu1_MenuClick(ByVal MenuNumber As Long)
    MsgBox MenuNumber
End Sub

Private Sub CtlVerticalMenu1_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    MsgBox MenuNumber & "  " & MenuItem
End Sub

