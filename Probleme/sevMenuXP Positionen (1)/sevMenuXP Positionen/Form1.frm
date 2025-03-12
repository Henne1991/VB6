VERSION 5.00
Object = "{C893915C-2B4D-4BCD-8D3C-4FB5197703F9}#1.10#0"; "sevMenuXP2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Neu sortieren"
      Height          =   1095
      Left            =   1050
      TabIndex        =   3
      Top             =   3060
      Width           =   4635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1. Menüeintrag entfernen und| neuen Eintrag an Position des 2. Eintrags anlegen"
      Height          =   1095
      Left            =   4800
      TabIndex        =   2
      Top             =   1170
      Width           =   4635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1. Menüeintrag entfernen und| neuen Eintrag an Position des 2. Eintrags anlegen"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1170
      Width           =   4635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Menü anzeigen mit Positionen"
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4635
   End
   Begin sevMenuXP.MenuBar MenuBar1 
      Left            =   1170
      Top             =   3060
      _ExtentX        =   794
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
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

Private mnuPopUp As sevMenuXP.MenuItem
Private mnuTemp As sevMenuXP.MenuItem
Private mnuPapierkorb As sevMenuXP.MenuItem


Private Sub Command1_Click()
   MenuBar1.PopUp mnuPopUp
End Sub

Private Sub Command2_Click()
   MenuItemEntfernen mnuPopUp, mnuPopUp.ChildItem(0)
End Sub

Private Sub MenuItemEntfernen(ByRef PopupItem As sevMenuXP.MenuItem, ByVal Item As sevMenuXP.MenuItem)
   
   Dim i As Integer
   Dim rs As Recordset
   
   Dim ItemIndex As Long
   Dim PopUpIndex As Long
   Dim PopUpKey As String
      
   ItemIndex = Item.Index
   PopUpIndex = PopupItem.Index
   PopUpKey = PopupItem.Key

   Set rs = New Recordset
   rs.Fields.Append "Index", adInteger
   rs.Fields.Append "Position", adInteger
   rs.Open

   For i = 0 To mnuPopUp.Children - 1
      rs.AddNew
      With mnuPopUp.ChildItem(i)
         rs!Index = .Index
         rs!Position = .Position
      End With
   Next

   rs.Filter = "Index<>" & ItemIndex
   Do Until rs.EOF
      MenuBar1.MenuItems.Move MenuBar1.MenuItems(CInt(rs!Index)), mnuTemp
      rs.MoveNext
   Loop

   MenuBar1.MenuItems.Remove ItemIndex
   MenuBar1.MenuItems.Remove PopUpKey
   
   Set mnuPopUp = MenuBar1.MenuItems.AddPopup()
   rs.Filter = "Index<>" & ItemIndex
   rs.Sort = "Position ASC"
   Do Until rs.EOF
      Set Item = MenuBar1.MenuItems(CInt(rs!Index))
      MenuBar1.MenuItems.Move Item, mnuPopUp, rs!Position + 100
      rs.MoveNext
   Loop
   MenuBar1.PopUp mnuPopUp
   
End Sub

Private Sub Command3_Click()
   Dim Item As sevMenuXP.MenuItem
   Set Item = mnuPopUp.ChildItem(0)
   Item.Visible = False
   Item.Key = ""
   Item.Position = -1000

      MenuBar1.MenuItems.Move Item, mnuTemp

   MenuBar1.MenuItems.Add("neu", , mnuPopUp, , , mnuPopUp.ChildItem(1).Position + 1).Tag = "neu"
   MenuBar1.PopUp mnuPopUp
End Sub

Private Sub Sortieren()
   Dim i As Integer
   Dim d
   
   While mnuPopUp.Children > 0
      MenuBar1.MenuItems.Move mnuPopUp.ChildItem(0), mnuTemp, mnuPopUp.ChildItem(0).Position
   Wend
      
   While mnuTemp.Children > 0
      MenuBar1.MenuItems.Move mnuTemp.ChildItem(0), mnuPopUp, mnuTemp.ChildItem(0).Position
   Wend
   
End Sub

Private Sub Command4_Click()
Sortieren
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim c As Control
   Set mnuPopUp = MenuBar1.MenuItems.AddPopup("mnuPopUp")
   Set mnuTemp = MenuBar1.MenuItems.AddPopup("mnuTemp")
Set mnuPapierkorb = MenuBar1.MenuItems.AddPopup("mnuPapierkorb")

   ChildrenLaden
   
   For Each c In Controls
      If TypeOf c Is CommandButton Then
         c.Caption = Replace(c.Caption, "|", vbNewLine)
      End If
   Next


End Sub

Private Sub ChildrenLaden()
Dim i As Integer
   For i = 1 To 10
      With MenuBar1.MenuItems.Add("Position " & i, "mnuTemp_" & i, mnuPopUp, , , i * 1000)
         .Tag = .Caption
      End With
   Next
End Sub

Private Sub MenuBar1_ItemPopup(ByVal Item As sevMenuXP.MenuItem)
   Dim i As Integer

   For i = 0 To Item.Children - 1
      Item.ChildItem(i).Caption = Item.ChildItem(i).Tag & " (Index: " & Item.ChildItem(i).Index & " >> Pos. " & Item.ChildItem(i).Position & ")"
   Next
End Sub
