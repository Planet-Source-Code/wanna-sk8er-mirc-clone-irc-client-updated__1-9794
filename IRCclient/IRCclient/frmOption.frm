VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Option"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2265
      Index           =   1
      Left            =   735
      TabIndex        =   37
      Top             =   3330
      Width           =   3105
      Begin VB.CheckBox chkIdentShow 
         Caption         =   "Show ident requests "
         Height          =   255
         Left            =   465
         TabIndex        =   42
         Top             =   1845
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtIdentUserID 
         Height          =   285
         Left            =   1080
         TabIndex        =   41
         Text            =   "sk8er"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtIdentSystem 
         Height          =   285
         Left            =   1080
         TabIndex        =   40
         Text            =   "UNIX"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtIdentPort 
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         Text            =   "113"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkIdent 
         Caption         =   "Enable Ident server"
         Height          =   255
         Left            =   465
         TabIndex        =   38
         Top             =   1575
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   45
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "System:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   44
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Port:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   43
         Top             =   1200
         Width           =   330
      End
   End
   Begin VB.ComboBox txtServer 
      Height          =   315
      Left            =   1485
      TabIndex        =   36
      Text            =   "irc.insiderz.net"
      Top             =   495
      Width           =   2085
   End
   Begin VB.TextBox txtRealName 
      Height          =   300
      Left            =   1815
      TabIndex        =   31
      Text            =   "Adam Wannamaker"
      Top             =   2850
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect to Server"
      Height          =   375
      Left            =   1620
      TabIndex        =   30
      Top             =   1050
      Width           =   1830
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1815
      TabIndex        =   29
      Text            =   "Sk8erCLONE"
      Top             =   1860
      Width           =   1830
   End
   Begin VB.TextBox txtMail 
      Height          =   285
      Left            =   1815
      TabIndex        =   28
      Text            =   "email@blah.net"
      Top             =   2325
      Width           =   1965
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   1
      Left            =   210
      ScaleHeight     =   3135
      ScaleWidth      =   3495
      TabIndex        =   0
      Top             =   5760
      Width           =   3495
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect to IRC Server"
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   570
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Invisible mode"
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtFullName 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cmbServers 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblStatic 
         Caption         =   "Alternative"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "Nickname:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "Email Address:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "Full Name:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "IRC Servers"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Index           =   0
      Left            =   3750
      ScaleHeight     =   3135
      ScaleWidth      =   3495
      TabIndex        =   13
      Top             =   5760
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2040
         TabIndex        =   27
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Okay"
         Height          =   315
         Left            =   480
         TabIndex        =   26
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Text            =   "6667"
         Top             =   2280
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "When connecting:"
         Height          =   1455
         Index           =   0
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   2415
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   720
            TabIndex        =   22
            Text            =   "99"
            Top             =   600
            Width           =   375
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Try next server in group"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   720
            TabIndex        =   18
            Text            =   "99"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblStatic 
            Caption         =   "second(s)"
            Height          =   255
            Index           =   8
            Left            =   1200
            TabIndex        =   24
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblStatic 
            Caption         =   "Delay:"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   23
            Top             =   660
            Width           =   495
         End
         Begin VB.Label lblStatic 
            Caption         =   "time(s)"
            Height          =   255
            Index           =   6
            Left            =   1200
            TabIndex        =   21
            Top             =   300
            Width           =   495
         End
         Begin VB.Label lblStatic 
            Caption         =   "Retry:"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Reconnect on disconnection"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Connect on start up"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblStatic 
         Caption         =   "Default port:"
         Height          =   255
         Index           =   9
         Left            =   960
         TabIndex        =   25
         Top             =   2325
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Full Name"
      Height          =   255
      Left            =   795
      TabIndex        =   35
      Top             =   2850
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Email Address"
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   34
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "NickName"
      Height          =   255
      Index           =   0
      Left            =   435
      TabIndex        =   33
      Top             =   1860
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "IRC Server"
      Height          =   255
      Index           =   0
      Left            =   285
      TabIndex        =   32
      Top             =   510
      Width           =   1095
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MyNick = txtNick
    Email = txtMail
    RealName = txtRealName
    
    If Len(txtServer) = 0 Then Exit Sub
    Connect txtServer, "6667"
    frmStatus.rtfStatus.SelColor = vbBlack
    frmStatus.rtfStatus.SelText = " *** Connecting to Server " & vbCrLf
IdentUserID = txtIdentUserID.Text
    If chkIdent Then
        On Error GoTo InUse
        frmMain.ident(0).Close
        If frmMain.ident(0).State <> sckListening Then
            frmMain.ident(0).LocalPort = Val(txtIdentPort)
            'mdiMain.ident(0).Bind Val(txtIdentPort), mdiMain.ident(0).LocalIP
            frmMain.ident(0).Listen
        End If
    End If
    
    'OK...unload form
    Unload Me
'Identd socket is used in another program - conflict
InUse:
    If Err.Number = 10048 Then
        MsgBox "Another program is using port " & frmMain.ident(0).LocalPort & "." & vbCrLf & "IdentD will be disabled." & vbCrLf & "You will need to close the other program and" & vbCrLf & "reopen to use Ident server.", vbOKOnly, "IRC IdentD problem"
        Resume Next
    End If
    
Unload Me
End Sub
