VERSION 5.00
Begin VB.Form frmDCCACCEPT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " IRC client DCC Chat"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   Icon            =   "frmDCCACCEPT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChat 
      Caption         =   "Chat!"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1290
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1905
      TabIndex        =   4
      Top             =   1290
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   150
      TabIndex        =   0
      Top             =   15
      Width           =   3255
      Begin VB.Label lblNickName 
         Alignment       =   2  'Center
         Caption         =   "NickName"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblIP 
         Alignment       =   2  'Center
         Caption         =   "IP"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblPort 
         Alignment       =   2  'Center
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmDCCACCEPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChat_Click()
    'CHAT_Index = CHAT_Index + 1
    'ReDim Preserve ChatWindow(CHAT_Index)
    'ReDim Preserve ChatWindowName(CHAT_Index)
    'Load ChatWindow(CHAT_Index)

    'ChatWindow(CHAT_Index).Caption = lblNickName
    'ChatWindowName(CHAT_Index) = lblNickName
    
    'Load mdiMain.Chat(CHAT_Index)
    'frmSocket.Chat.Connect lblIP, lblPort
    'Unload Me
End Sub

Private Sub Form_Load()
Beep
End Sub
