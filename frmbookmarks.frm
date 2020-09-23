VERSION 5.00
Begin VB.Form frmbookmarks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bookmarks"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBookmarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Site to add"
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add Current Site"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   4455
   End
   Begin VB.ListBox lstBookmarks 
      Height          =   2010
      ItemData        =   "frmbookmarks.frx":0000
      Left            =   120
      List            =   "frmbookmarks.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "frmbookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I got this code from another submission on PSC but i cant find it anymore...

Private Sub cmdAdd_Click()
Dim strBookmark As String
Open App.Path & "\Bookmarks.txt" For Append As #1
Print #1, frmmain.combo1.Text
lstBookmarks.AddItem frmmain.combo1.Text
Close #1
MsgBox "The combo1 has been saved to your bookmark list.", 8, ""
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub Form_Load()


    On Error GoTo FileError
    Open App.Path & "\Bookmarks.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, a$
        lstBookmarks.AddItem a$
    Loop
    Close #1
FileError:
    Open App.Path & "\Bookmarks.list" For Output As #1
    Close #1
End Sub



Private Sub lstBookmarks_DblClick()
    If InStr(1, "lstBookmarks.text", "Http:\\") = 0 Then
        frmmain.wb.Navigate lstBookmarks.Text
    Else
        frmmain.wb.Navigate "Http:\\" & lstBookmarks.Text
    End If
End Sub


Private Sub txtBookmarks_KeyPress(KeyAscii As Integer)
Dim strBookamrk As String
    If KeyAscii = 13 Then
        Open App.Path & "\Bookmarks.txt" For Append As #1
        Print #1, txtBookmarks.Text
            lstBookmarks.AddItem txtBookmarks.Text
            
             txtBookmarks.Text = strBookamrk
        Close #1
            frmbookmarks.Visible = False
            MsgBox "The combo1 has been saved to your bookmark list.", 8, ""
    End If

End Sub
