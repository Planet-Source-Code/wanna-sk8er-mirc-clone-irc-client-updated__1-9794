VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "IRC Client"
   ClientHeight    =   7545
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList imgChannel 
      Left            =   2670
      Top             =   2115
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2066
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2482
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":281E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3372
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":370E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":457E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "connect"
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "option"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "folder"
            Object.ToolTipText     =   "Channels folder"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "list"
            Object.ToolTipText     =   "List channels"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "color"
            Object.ToolTipText     =   "Colors"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "address"
            Object.ToolTipText     =   "Address Book"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dccsend"
            Object.ToolTipText     =   "DCC Send File"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dccchat"
            Object.ToolTipText     =   "DCC Chat"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tile"
            Object.ToolTipText     =   "Tile Windows"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cascade"
            Object.ToolTipText     =   "Cascade Windows"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "about"
            Object.ToolTipText     =   "About"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock ident 
      Index           =   0
      Left            =   3465
      Top             =   2385
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuReConnect 
         Caption         =   "Reconnect..."
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu hy00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOption 
         Caption         =   "Options"
      End
      Begin VB.Menu hy03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearStatus 
         Caption         =   "Clear Status..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTile 
         Caption         =   "&Tile"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu hy02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAuto 
         Caption         =   "Auto"
         Begin VB.Menu mnuAutoTile 
            Caption         =   "Tile"
         End
         Begin VB.Menu mnuAutoCascade 
            Caption         =   "Cascade"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuBrowser 
         Caption         =   "Online Browser"
      End
      Begin VB.Menu hy04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About IRC client"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ident_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    
    Dim socket As Variant
    For Each socket In ident
        If (socket.State = sckClosed Or socket.State = sckError) Then
            socket.Close
            socket.Accept requestID
            'frmOptions.txtIdentLog = frmOptions.txtIdentLog & socket.RemoteHostIP & "[" & requestID & "]" & vbCrLf
            socket.SendData socket.LocalPort & ", " & requestID & ":USERID:WIN32:" & IdentUserID & vbCrLf
            '1236, 7000 : USERID : UNIX : higher
            Call DoColor(frmStatus.rtfStatus, "6* Ident request from " & frmMain.ident(0).RemoteHostIP)
            Call DoColor(frmStatus.rtfStatus, "6* Ident Reply: " & socket.LocalPort & ", " & requestID & ":USERID:WIN32:" & IdentUserID)
            frmStatus.rtfStatus.SelText = "-" & vbCrLf
            DoEvents
            socket.Close
            Exit For
        End If
    Next socket
End Sub



Private Sub MDIForm_Load()
    LoadColor
    Call AddTaskbar("Status", 1)
        
    Dim i As Integer
    Const maxtcp = 10
    For i = 1 To maxtcp
        Load ident(i)
    Next i
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call mnuExit_Click
Dim i As Integer
Const maxtcp = 10
    For i = 1 To maxtcp
        Unload ident(i)
    Next i
End Sub

Private Sub MDIForm_Resize()
If mnuAutoCascade.Checked = True Then
    frmMain.Arrange vbCascade
    Else
End If

If mnuAutoTile.Checked = True Then
    frmMain.Arrange vbTileHorizontal
    Else
End If
End Sub

Private Sub mnuArrangeIcons_Click()
    frmMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuAutoCascade_Click()
If mnuAutoCascade.Checked = False Then
        mnuAutoCascade.Checked = True
    Else
        mnuAutoCascade.Checked = False
    End If
If mnuAutoTile.Checked = True Then
    mnuAutoTile.Checked = False
    Else
End If
End Sub

Private Sub mnuAutoTile_Click()
If mnuAutoTile.Checked = False Then
        mnuAutoTile.Checked = True
     Else
        mnuAutoTile.Checked = False
    End If
If mnuAutoCascade.Checked = True Then
    mnuAutoCascade.Checked = False
    Else
End If
End Sub

Private Sub mnuCascade_Click()
    frmMain.Arrange vbCascade
End Sub

Private Sub mnuConnect_Click()
frmOption.Command1.Value = True
End Sub

Private Sub mnuDisconnect_Click()
    If frmSocket.socket.State = sckConnected Then
        SendData "QUIT"
    End If
    Timeout 0.5
    frmSocket.socket.Close
    LogText frmStatus.rtfStatus, "*** Disconnected", udtColor.colorQuit
End Sub

Private Sub mnuExit_Click()
    Disconnect
    End
End Sub

Private Sub mnuOption_Click()
    frmOption.Show 1
End Sub

Private Sub mnuReConnect_Click()
frmOption.Command1.Value = True
End Sub

Private Sub mnuTile_Click()
    frmMain.Arrange vbTileHorizontal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Set up the toolbar
Select Case Button.Key
Case "connect"
        If Toolbar1.Buttons(1).Image = 1 Then
            Toolbar1.Buttons(1).Image = 2
                Toolbar1.Buttons(1).ToolTipText = "Disconnect"
                    mnuConnect_Click
                Else
                    Toolbar1.Buttons(1).Image = 1
                        Toolbar1.Buttons(1).ToolTipText = "Connect"
                            mnuDisconnect_Click
                        End If
Case "option"
mnuOption_Click
Case "folder"
MsgBox "Not done yet.."
Case "list"
MsgBox "Not done yet.."
Case "color"
frmColor.Show
Case "address"
MsgBox "Not done yet.."
Case "dccsend"
MsgBox "Not done yet.."
Case "dccchat"
MsgBox "Not done yet.."
Case "tile"
mnuTile_Click
Case "cascade"
mnuCascade_Click
Case "help"
MsgBox "Not done yet.."
Case "about"
MsgBox "Not done yet.."
End Select
End Sub
