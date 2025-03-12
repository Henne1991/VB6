VERSION 5.00
Object = "{C893915C-2B4D-4BCD-8D3C-4FB5197703F9}#1.10#0"; "sevMenuXP2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "MsgBox"
      Height          =   705
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin sevMenuXP.MenuBar MenuBar1 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   794
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighLiteMode    =   3
      IconHighLiteMode=   3
      TextHighLiteColor=   16711680
      TextHotLightColor=   255
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Version2        =   -1  'True
   End
   Begin VB.Menu MenuDummy 
      Caption         =   "sevMenuXP"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   MsgBox "Hallo"
End Sub

Private Sub Form_Load()
   Dim m As MenuItem
   Set m = MenuBar1.MenuItems.Add("Datei")
   MenuBar1.MenuItems.Add "Hotkey F2", "F2", m, , , , , , , vbKeyF2
End Sub

Private Sub MenuBar1_ItemClick(ByVal Item As sevMenuXP.MenuItem)
   If Item.Key = "F2" Then MsgBox "F2 wurde gedrückt"
End Sub
