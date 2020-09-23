VERSION 5.00
Object = "{1339B53E-3453-11D2-93B9-000000000000}#1.0#0"; "mozctl.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "MozBrowser"
   ClientHeight    =   2865
   ClientLeft      =   165
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Text            =   "URL"
      Top             =   840
      Width           =   4695
   End
   Begin MOZILLACONTROLLibCtl.MozillaBrowser wb 
      Height          =   1455
      Left            =   0
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2610
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1402
      ButtonWidth     =   1217
      ButtonHeight    =   1402
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Key             =   "home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   31
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0024
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2486
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6E35
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9451
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu print 
         Caption         =   "Print"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "Options"
      Begin VB.Menu block 
         Caption         =   "Block Popups"
      End
      Begin VB.Menu nwindow 
         Caption         =   "New Window"
      End
      Begin VB.Menu favorites 
         Caption         =   "Bookmarks"
         Begin VB.Menu view 
            Caption         =   "View"
         End
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub block_Click()
If block.Checked = True Then
block.Checked = False
Else
block.Checked = True
End If
End Sub



Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
wb.GoHome
Me.WindowState = 2
End Sub

Private Sub Form_Resize()
wb.Height = Me.Height - 2175
wb.Width = Me.Width - 100
combo1.Width = Me.Width - 100
End Sub

Private Sub nwindow_Click()
Dim frm As frmmain
Set frm = New frmmain
Set ppDisp = frmmain.wb.object
frm.Show
frm.wb.Navigate "about:blank"
End Sub

Private Sub print_Click()
Print wb.Document
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key
    Case "back"
        wb.GoBack
    Case "forward"
        wb.GoForward
    Case "Stop"
        wb.Stop
    Case "refresh"
        wb.Refresh
    Case "home"
        wb.GoHome
    Case "Search"
    End Select
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
wb.Navigate combo1.Text
combo1.AddItem combo1.Text
End If
End Sub


Private Sub view_Click()
frmbookmarks.Show
End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, url As Variant)
combo1.Text = wb.LocationURL
Me.Caption = wb.LocationName
End Sub

Private Sub wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
If block.Checked = True Then
Cancel = True
Else
Dim frm As frmmain
Set frm = New frmmain
Set ppDisp = frmmain.wb.object
frm.Show
End If
End Sub

Private Sub wb_StatusTextChange(ByVal Text As String)
StatusBar1.SimpleText = Text
End Sub
