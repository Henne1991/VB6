VERSION 5.00
Object = "{34A21D19-E940-4615-8449-D99CB165CB63}#1.1#0"; "sevDataGrid3.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19260
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   19260
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   2565
      Left            =   3840
      TabIndex        =   1
      Top             =   1860
      Width           =   3825
   End
   Begin sevDataGrid3.sevGrid sevGrid1 
      Height          =   10155
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19245
      _ExtentX        =   33946
      _ExtentY        =   17912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ColumnHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FilterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoNumberFixedCol=   -1  'True
      FixedCol        =   -1  'True
      Version3        =   -1  'True
      FilterBorderColor=   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   sevGrid1.RefreshEx
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_Resize()
   sevGrid1.Move 0, 0, Width, Height
End Sub

Private Sub sevGrid1_AfterAddNew()

End Sub
