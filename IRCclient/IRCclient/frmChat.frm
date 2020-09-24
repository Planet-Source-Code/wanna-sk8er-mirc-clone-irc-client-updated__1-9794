VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtDCC 
      Height          =   2490
      Left            =   45
      TabIndex        =   1
      Top             =   15
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   4392
      _Version        =   393217
      BorderStyle     =   0
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   2910
      Width           =   4590
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
    On Error Resume Next
    txtSend.Top = txtDCC.Height + 10
    txtDCC.Width = Me.ScaleWidth
    txtSend.Width = Me.ScaleWidth
    txtDCC.Height = (Me.Height - txtSend.Height - 400)
End Sub

Private Sub txtDCC_Change()
txtDCC.SelStart = Len(rtfStatus.Text)
End Sub

