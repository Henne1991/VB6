VERSION 5.00
Object = "{34A21D19-E940-4615-8449-D99CB165CB63}#1.1#0"; "sevDataGrid3.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   30
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4440
      Width           =   6765
   End
   Begin sevDataGrid3.sevGrid sevGrid1 
      Height          =   4335
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7646
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
      CellSpacingH    =   1
      CellSpacingV    =   1
      Version3        =   -1  'True
      FilterBorderColor=   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Dim i As Integer
   Dim rs As New ADODB.Recordset

   For i = 1 To 10
      sevGrid1.ColumnAdd , "Spalte " & i
      rs.Fields.Append "Spalte " & i, adVarChar, 500
   Next

   rs.Open
   For i = 1 To 100
      rs.AddNew
      rs("Spalte 1").Value = i

   Next
   Set sevGrid1.Recordset = rs.Clone
   sevGrid1.Refresh
   Aufruf = 0
End Sub

Private Sub sevGrid1_Scroll(ByVal nScrollBar As sevDataGrid3.sevGridScrollBar)
   Aufruf = Aufruf + 1
End Sub

Private Property Get Aufruf() As Long
   Aufruf = Val(sevGrid1.Tag)
End Property

Private Property Let Aufruf(n As Long)
   sevGrid1.Tag = n
   Text1.Text = "Event-Aufrufe: " & n
End Property
