VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTunnel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tunnel"
   ClientHeight    =   5310
   ClientLeft      =   2115
   ClientTop       =   1545
   ClientWidth     =   3855
   Icon            =   "frmTunnel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   3855
   Begin VB.PictureBox picPaused 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   435
      Left            =   880
      Picture         =   "frmTunnel.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   2070
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   2070
   End
   Begin MSComctlLib.StatusBar stbHighScore 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5055
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "High Score:"
            TextSave        =   "High Score:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   4604
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer timCrashmessage 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3360
      Top             =   0
   End
   Begin VB.Frame fraYoucrashed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   80
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   3690
      Begin VB.PictureBox picYoucrashed 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   0
         Left            =   0
         Picture         =   "frmTunnel.frx":326C
         ScaleHeight     =   435
         ScaleWidth      =   3690
         TabIndex        =   4
         Top             =   0
         Width           =   3690
      End
      Begin VB.PictureBox picYoucrashed 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   1
         Left            =   0
         Picture         =   "frmTunnel.frx":8682
         ScaleHeight     =   435
         ScaleWidth      =   3690
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.PictureBox picYoucrashed 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   2
         Left            =   0
         Picture         =   "frmTunnel.frx":DA98
         ScaleHeight     =   435
         ScaleWidth      =   3690
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   3690
      End
   End
   Begin VB.Timer timMainTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.StatusBar stbScore 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Score:"
            TextSave        =   "Score:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   4604
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
      MousePointer    =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraShip 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   1600
      MousePointer    =   12  'No Drop
      TabIndex        =   6
      Top             =   4140
      Width           =   630
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   0
         Left            =   0
         Picture         =   "frmTunnel.frx":12EAE
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   8
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   1
         Left            =   0
         Picture         =   "frmTunnel.frx":143F0
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   7
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   2
         Left            =   0
         Picture         =   "frmTunnel.frx":15932
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   11
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   3
         Left            =   0
         Picture         =   "frmTunnel.frx":15DA4
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   12
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   4
         Left            =   0
         Picture         =   "frmTunnel.frx":1691E
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   13
         Top             =   0
         Width           =   630
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   5
         Left            =   0
         Picture         =   "frmTunnel.frx":17498
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   6
         Left            =   0
         Picture         =   "frmTunnel.frx":1790A
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox picShip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   630
         Index           =   7
         Left            =   0
         Picture         =   "frmTunnel.frx":17D7C
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   19
      Left            =   3100
      Top             =   4560
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   18
      Left            =   3100
      Top             =   4320
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   17
      Left            =   3100
      Top             =   4080
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   16
      Left            =   3100
      Top             =   3840
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   15
      Left            =   3100
      Top             =   3600
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   14
      Left            =   3100
      Top             =   3360
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   13
      Left            =   3100
      Top             =   3120
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   12
      Left            =   3100
      Top             =   2880
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   11
      Left            =   3100
      Top             =   2640
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   10
      Left            =   3100
      Top             =   2400
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   9
      Left            =   3100
      Top             =   2160
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   8
      Left            =   3100
      Top             =   1920
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   7
      Left            =   3100
      Top             =   1680
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   6
      Left            =   3100
      Top             =   1440
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   5
      Left            =   3100
      Top             =   1200
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   4
      Left            =   3100
      Top             =   960
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   3
      Left            =   3100
      Top             =   720
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   2
      Left            =   3100
      Top             =   480
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   1
      Left            =   3100
      Top             =   240
      Width           =   60
   End
   Begin VB.Shape shpRightEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   0
      Left            =   3100
      Top             =   0
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   19
      Left            =   700
      Top             =   4560
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   18
      Left            =   700
      Top             =   4320
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   17
      Left            =   700
      Top             =   4080
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   16
      Left            =   700
      Top             =   3840
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   15
      Left            =   700
      Top             =   3600
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   14
      Left            =   700
      Top             =   3360
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   13
      Left            =   700
      Top             =   3120
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   12
      Left            =   700
      Top             =   2880
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   11
      Left            =   700
      Top             =   2640
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   10
      Left            =   700
      Top             =   2400
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   9
      Left            =   700
      Top             =   2160
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   8
      Left            =   700
      Top             =   1920
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   7
      Left            =   700
      Top             =   1680
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   6
      Left            =   700
      Top             =   1440
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   5
      Left            =   700
      Top             =   1200
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   4
      Left            =   700
      Top             =   960
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   3
      Left            =   700
      Top             =   720
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   2
      Left            =   700
      Top             =   480
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   1
      Left            =   700
      Top             =   240
      Width           =   60
   End
   Begin VB.Shape shpLeftEdge 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   0
      Left            =   700
      Top             =   0
      Width           =   60
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewgame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFilePause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditShip 
         Caption         =   "&Ship"
         Begin VB.Menu mnuEditShipShip 
            Caption         =   "Classic &Flyer"
            Index           =   0
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuEditShipShip 
            Caption         =   "Double &V-Flyer"
            Index           =   1
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuEditShipShip 
            Caption         =   "&Racing Car"
            Index           =   2
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuEditShipShip 
            Caption         =   "&Tank"
            Index           =   3
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnuEditShipShip 
            Caption         =   "&W-Flyer"
            Index           =   4
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mnuEditShipShip 
            Caption         =   "&X-Flyer"
            Index           =   5
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu mnuEditShipShip 
            Caption         =   "X-Flyer &LE"
            Index           =   6
            Shortcut        =   ^{F7}
         End
      End
      Begin VB.Menu mnuEditColor 
         Caption         =   "Tunnel Edge &Color"
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&Black"
            Index           =   0
            Shortcut        =   +{F1}
         End
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&White"
            Index           =   1
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&Blue"
            Index           =   2
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&Cyan"
            Index           =   3
            Shortcut        =   +{F4}
         End
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&Green"
            Index           =   4
            Shortcut        =   +{F5}
         End
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&Magenta"
            Index           =   5
            Shortcut        =   +{F6}
         End
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&Red"
            Index           =   6
            Shortcut        =   +{F7}
         End
         Begin VB.Menu mnuEditColorColor 
            Caption         =   "&Yellow"
            Index           =   7
            Shortcut        =   +{F8}
         End
      End
      Begin VB.Menu mnuEditChangespeed 
         Caption         =   "&Change Speed"
      End
      Begin VB.Menu mnuEditResethighscore 
         Caption         =   "&Reset High Score"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpKeys 
         Caption         =   "&Keys"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmTunnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileNumber As Integer
Dim FormWidth As Integer
Dim i As Integer
Dim j As Integer
Dim Index As Integer
Dim Message As String
Dim Paused As Boolean
Dim ShipCrashed As Boolean
Dim TheColor As String

Const ExplodingShip = 7
Const NumberofColors = 7
Const NumberofShips = 6

Const Black = 0
Const White = 16777215
Const Blue = 16711680
Const Cyan = 16776960
Const Green = 65280
Const Magenta = 16711935
Const Red = 255
Const Yellow = 65535
'**********************************************************
'**********************************************************
'TUNNEL (Version 1.0)
'Programmed by: Anton Venema
'Date: September 9, 2000
'
'Tunnel is a game in which a ship or object moves through
'a randomly shrinking/expanding "tunnel" while trying not
'to hit the edges.
'
'You can choose between two ships - the default X-Ship or
'the X-Ship Compact.  The X-Ship Compact is basically the
'same as the X-Ship except that it lacks the extra border
'that connects the four points of the "X" shape.
'
'You can also choose the background color to the default
'white, as well as blue, cyan, green, magenta, red, and
'yellow.
'
'The keys for the game are listed in the Keys section of
'the Help menu.
'
'Pressing [Ctrl] + [N] will start a new game.  Pressing F3
'will pause/unpause a game in progress.
'
'The current score is displayed in the top status bar.  The
'high score is displayed int he bottowm status bar.
'
'Coding Copyright © 2000
'This code may not be freely used in personal programs.  It
'is intended to be used for educational/informational
'purposes only.
'**********************************************************
'**********************************************************

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If ShipCrashed = False Then

        'Left arrow moves 120 twips
        'Shift, Left arrow moves 60 twips
        Select Case KeyCode
            Case vbKeyLeft
                Select Case Shift
                    Case 0
                        fraShip.Left = fraShip.Left - 120
                    Case 1
                        fraShip.Left = fraShip.Left - 60
                End Select
            Case vbKeyRight
                Select Case Shift
                    Case 0
                        fraShip.Left = fraShip.Left + 120
                    Case 1
                        fraShip.Left = fraShip.Left + 60
                End Select
            Case Else
                Exit Sub
        End Select

        'After moving, check for crash
        CrashDetect
    End If

End Sub

Private Sub FirstSetup()
Dim Shift As Integer
Dim TotalShift As Integer

    'Reset current score
    stbScore.Panels(2).Text = "0"

    'Reset Pause and Change Ship menu items
    mnuFilePause.Enabled = True
    mnuFilePause.Caption = "&Pause"
    Paused = False
    picPaused.Visible = False
    mnuEditShip.Enabled = True

    'Hide crashed message
    fraYoucrashed.Visible = False
    timCrashmessage.Enabled = False

    Reset

    'Set intial shift amount to 0
    TotalShift = 0

    'Set random scheme for tunnel edges
    For i = 18 To 0 Step -1
        Randomize
        Shift = CInt(Rnd)

        'Set shifting amount
        Select Case Shift

            'Move left 60 twips if possible
            Case 0
                If shpLeftEdge(i).Left + TotalShift - 60 < 0 Or shpRightEdge(i).Left + TotalShift - 60 < 0 Then
                    TotalShift = TotalShift + 60
                Else
                    TotalShift = TotalShift - 60
                End If

            'Move right 60 twips if possible
            Case 1
                If shpRightEdge(i).Left + TotalShift + 60 > (frmTunnel.Width - shpRightEdge(i).Width) Or shpLeftEdge(i).Left + TotalShift + 60 > (frmTunnel.Width - shpLeftEdge(i).Width) Then
                    TotalShift = TotalShift - 60
                Else
                    TotalShift = TotalShift + 60
                End If
        End Select

        'Move edges of tunnel
        shpLeftEdge(i).Left = shpLeftEdge(i).Left + TotalShift
        shpRightEdge(i).Left = shpRightEdge(i).Left + TotalShift
    Next i

End Sub

Private Sub Reset()

    'Straighten left and right edges of tunnel
    For i = 19 To 0 Step -1
        shpLeftEdge(i).Left = (frmTunnel.Width / 2) - 1200
        shpRightEdge(i).Left = (frmTunnel.Width / 2) + 1200
    Next i

    'Center ship in screen
    fraShip.Left = (frmTunnel.Width / 2) - (fraShip.Width / 2)

    'Hide crashed ship
    picShip(ExplodingShip).Visible = False

    'Reset the ship that was used
    LoadShip

    'Ship has not crashed yet
    ShipCrashed = False

End Sub

Private Sub Form_Load()

    If Screen.Width < 4000 Then
        Message = MsgBox("Your current screen size is not large enough to support Tunnel.  The minimum resolution for running Tunnel is: 266.6 x 200", , "Screen Size")
        End
    End If

    'Set Paused value and run intial setup
    'routines
    Paused = True

    LoadColorScheme

    LoadShip

    LoadSpeed

    SetHighScoreOnBar

    'Set form width
    frmTunnel.Width = Screen.Width - 1000

    'Center form on screen
    frmTunnel.Left = (Screen.Width / 2) - (frmTunnel.Width / 2)

    'Center ship on form
    fraShip.Left = (frmTunnel.Width / 2) - (fraShip.Width / 2)

    'Adjust frame to picture size
    fraShip.Width = picShip(0).Width
    fraShip.Height = picShip(0).Width

    'Properly align tunnel edges
    For i = 0 To 19
        shpLeftEdge(i).Left = (frmTunnel.Width / 2) - 1250
        shpRightEdge(i).Left = (frmTunnel.Width / 2) + 1250
    Next i

    'Set crash message properties
    With fraYoucrashed
        .Width = picYoucrashed(0).Width
        .Height = picYoucrashed(0).Height
        .Left = (frmTunnel.Width / 2) - (fraYoucrashed.Width / 2)
        .Top = (frmTunnel.Height / 2) - 1000
    End With
    
    'Set paused message properties
    With picPaused
        .Left = (frmTunnel.Width / 2) - (picPaused.Width / 2)
        .Top = (frmTunnel.Height / 2) - 1000
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Check to see if high score has been reached a game is
    'currently in progress
    If mnuFilePause.Enabled = True Then
        CheckForHighScore
    End If

    'Exit program
    End

End Sub

Private Sub mnuEditChangespeed_Click()
Dim TheSpeed As String

    'Pause game
    If mnuFilePause.Enabled = True Then
        Paused = False
        mnuFilePause_Click
    End If

InputSpeed:

    'Input the speed (in milliseconds)
    TheSpeed = InputBox("Enter the number of milliseconds (between 1 and 1000) that you wish to have between the move of each edge block.  (A lower number results in a faster scrolling speed.)" + Chr(13) + Chr(13) + "Current speed: " + CStr(timMainTimer.Interval) + " milliseconds", "Change Speed", CStr(timMainTimer.Interval))

    'If cancel was pressed, exit the subroutine
    If TheSpeed = "" Then
        Exit Sub
    End If

    'Check to make sure that only digits are contained in the entry
    For i = 1 To Len(TheSpeed)
        If Not (Asc(Mid(TheSpeed, i, 1)) >= Asc("0") And Asc(Mid(TheSpeed, i, 1)) <= Asc("9")) Then
            GoTo InputSpeed
        End If
    Next i

    'Check to make sure the number is between 1 and 1000
    If Not (CDbl(TheSpeed) >= 1 And CDbl(TheSpeed) <= 1000) Then
        GoTo InputSpeed
    End If

    'Set timer interval
    timMainTimer.Interval = CInt(TheSpeed)

End Sub

Private Sub mnuEditColorColor_Click(Index As Integer)

    'Uncheck all menu items
    For i = 0 To 7
        mnuEditColorColor(i).Checked = False
    Next i

    'Check clicked menu item
    mnuEditColorColor(Index).Checked = True

    'Set background color
    Select Case Index
        Case 0
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbBlack
                shpRightEdge(i).FillColor = vbBlack
            Next i
        Case 1
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbWhite
                shpRightEdge(i).FillColor = vbWhite
            Next i
        Case 2
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbBlue
                shpRightEdge(i).FillColor = vbBlue
            Next i
        Case 3
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbCyan
                shpRightEdge(i).FillColor = vbCyan
            Next i
        Case 4
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbGreen
                shpRightEdge(i).FillColor = vbGreen
            Next i
        Case 5
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbMagenta
                shpRightEdge(i).FillColor = vbMagenta
            Next i
        Case 6
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbRed
                shpRightEdge(i).FillColor = vbRed
            Next i
        Case 7
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbYellow
                shpRightEdge(i).FillColor = vbYellow
            Next i
    End Select

    'Save background color to file
    FileNumber = FreeFile

    Open "tcolor.tdf" For Output As #FileNumber
        Write #FileNumber, shpLeftEdge(0).FillColor
    Close #FileNumber

End Sub

Private Sub mnuEditResethighscore_Click()

    'Pause game
    If mnuFilePause.Enabled = True Then
        Paused = False
        mnuFilePause_Click
    End If

    'Make sure that the user desires to reset the high
    'score data
    Message = MsgBox("Are you sure you wish to reset the high score?", vbYesNo, "Reset High Score")

    'If the user does not wish to reset the high score
    'data after all, then exit the routine
    If Message = vbNo Then
        Exit Sub
    End If

    'Reset the high score data
    CreateDefaultHighScore

    'Display the new data on panel
    SetHighScoreOnBar

End Sub

Private Sub mnuEditShipShip_Click(Index As Integer)

    'Uncheck all menu items
    For i = 0 To NumberofShips
        mnuEditShipShip(i).Checked = False
    Next i

    'Check clicked menu item
    mnuEditShipShip(Index).Checked = True

    'Hide all ships
    For i = 0 To NumberofShips
        picShip(i).Visible = False
    Next i

    'Show ship that corresponds to clicked menu item
    picShip(Index).Visible = True

    'Save ship number to file
    FileNumber = FreeFile

    Open "tship.tdf" For Output As #FileNumber
        Write #FileNumber, Index
    Close #FileNumber

End Sub

Private Sub mnuFileExit_Click()

    'Exit the program
    End

End Sub

Private Sub mnuFileNewgame_Click()

    'If a game is currently in progress, check to make sure
    ' the user wishes to quit and lose his/her current game
    If mnuFilePause.Enabled = True Then
        Paused = False
        mnuFilePause_Click
        Message = MsgBox("Are you sure you wish to quit your current game?", vbYesNo, "New Game")
        If Message = vbNo Then
            Exit Sub
        End If
    End If

    'Run initial setup
    FirstSetup

    'Start the main timer that runs the game if not paused
    If Paused = False Then
        timMainTimer.Enabled = True
    End If

End Sub

Private Sub mnuFilePause_Click()

    '*************************************
    '* Routine that handles a Pause menu *
    '*                                   *
    '* Pause timer, set paused value,    *
    '* and change caption of menu item   *
    '*************************************

    If Paused = False Then
        timMainTimer.Enabled = False
        Paused = True
        mnuFilePause.Caption = "Un&pause"
        picPaused.Visible = True
        Exit Sub
    ElseIf Paused = True Then
        timMainTimer.Enabled = True
        Paused = False
        mnuFilePause.Caption = "&Pause"
        picPaused.Visible = False
        Exit Sub
    End If

End Sub

Private Sub mnuHelpAbout_Click()

    'Pause game
    If mnuFilePause.Enabled = True Then
        Paused = False
        mnuFilePause_Click
    End If

    'Display About screen
    frmAbout.Show vbModal

End Sub

Private Sub mnuHelpKeys_Click()

    'Pause game
    If mnuFilePause.Enabled = True Then
        Paused = False
        mnuFilePause_Click
    End If

    'Display dialog box explaining the keys
    Message = MsgBox("[Left Arrow] - Move left (120 twips)" + Chr(13) + "[Right Arrow] - Move right (120 twips)" + Chr(13) + "[Shift + Left Arrow] - Move left (60 twips)" + Chr(13) + "[Shift + Right Arrow] - Move right (60 twips)", , "Tunnel Keys")

End Sub

Private Sub timCrashmessage_Timer()
Dim VisiblePicture As Integer

    '*************************************************
    '* Routine that rotates crash message (3 frames) *
    '*************************************************

    'Get index of currently displayed frame
    For i = 0 To 2
        If picYoucrashed(i).Visible = True Then
            VisiblePicture = i
        End If
    Next i

    'Hide currently displayed frame
    picYoucrashed(VisiblePicture).Visible = False

    'Get index of next displayed frame
    VisiblePicture = VisiblePicture + 1
    If VisiblePicture = 3 Then
        VisiblePicture = 0
    End If

    'Display next frame
    picYoucrashed(VisiblePicture).Visible = True

End Sub

Private Sub timMainTimer_Timer()

    'Increase score by 1
    stbScore.Panels(2).Text = CStr(CDbl(stbScore.Panels(2).Text) + 1)

    'Give impression of tunnel moving
    Scroll

End Sub

Private Sub Scroll()
Dim RandomNumber As Integer
Dim TwipMove As Integer

    '*********************************
    '* Tunnel Edge Scrolling Routine *
    '*********************************

    'Starting with the bottom block, take the position of
    'the above block and apply it to the current block,
    'randomly shrinking or increasing the size of the tunnel
    'by ten twips every time.
    If shpRightEdge(19).Left - shpLeftEdge(19).Left > (fraShip.Width + 500) And shpLeftEdge(19).Left > 0 And shpRightEdge(19).Left < frmTunnel.Width Then
        Randomize
        RandomNumber = CInt(Rnd * 10)
        Select Case RandomNumber
            Case 0 Or 1 Or 2 Or 3 Or 4
                TwipMove = 10
            Case 5 Or 6 Or 7 Or 8 Or 9 Or 10
                TwipMove = (-10)
        End Select
    Else
        TwipMove = 0
    End If

    For i = 19 To 1 Step -1
        shpLeftEdge(i).Left = shpLeftEdge(i - 1).Left + TwipMove
        shpRightEdge(i).Left = shpRightEdge(i - 1).Left - TwipMove

        'For possible touching blocks, check to see if the
        'scrolling caused a block to touch the ship
        If i >= 17 And i <= 19 Then
            CrashDetect

            'If a block touched the ship, exit the routine
            If ShipCrashed = True Then
                Exit Sub
            End If
        End If
    Next i

    'Create random number (0 or 1)
    Randomize
    RandomNumber = CInt(Rnd)

    'Set shift amount for upper block (appears as new block)
    Select Case RandomNumber
        Case 0
            If (shpLeftEdge(0).Left - shpLeftEdge(0).Width) < 0 Or (shpRightEdge(0).Left - shpRightEdge(0).Width) < 0 Then
                Shift = 60
            Else
                Shift = -60
            End If
        Case 1
            If (shpRightEdge(0).Left + shpRightEdge(0).Width) > (frmTunnel.Width - 180) Or (shpLeftEdge(0).Left + shpLeftEdge(0).Width) > (frmTunnel.Width - 180) Then
                Shift = -60
            Else
                Shift = 60
            End If
    End Select

    'Take the second block's position and shift it left or
    'right and apply it to the top block
    shpLeftEdge(0).Left = shpLeftEdge(1).Left + Shift
    shpRightEdge(0).Left = shpRightEdge(1).Left + Shift

End Sub

Private Sub CrashDetect()

    '***************************
    '* Crash Detection Routine *
    '***************************

    'Check to see if the ship overlaps any bricks
    For j = 17 To 19
        If fraShip.Left < (shpLeftEdge(j).Left + shpLeftEdge(j).Width) Or fraShip.Left > (shpRightEdge(j).Left - fraShip.Width) Then

            'If there is an overlap then call crash routine,
            'set ShipCrashed value, and exit routine
            Crash
            ShipCrashed = True
            Exit Sub
        End If
    Next j

End Sub

Private Sub Crash()

    'Halt main timer, disable Pause menu item, disable
    'Change Ship menu item, hide tunnel edges, hide ship,
    'display frame containing crash message, start crash
    'message timer (rotates the message frames), and
    'display explosion picture
    timMainTimer.Enabled = False

    mnuFilePause.Enabled = False

    mnuEditShip.Enabled = False

    fraYoucrashed.Visible = True

    timCrashmessage.Enabled = True

    For i = 0 To NumberofShips
        picShip(i).Visible = False
    Next i

    picShip(ExplodingShip).Visible = True

    'Check to see if a new high score was reached
    CheckForHighScore

End Sub

Private Sub LoadColorScheme()
On Error GoTo ErrorHandler

    'Extract the saved color from file
    FileNumber = FreeFile

    Open "tcolor.tdf" For Input As #FileNumber
        Input #FileNumber, TheColor
    Close #FileNumber

    'Set background's color corresponding to file and set
    'Index value for menu checking
    Select Case TheColor
        Case Black
            'Black
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbBlack
                shpRightEdge(i).FillColor = vbBlack
            Next i
            Index = 0
        Case White
            'White
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbWhite
                shpRightEdge(i).FillColor = vbWhite
            Next i
            Index = 1
        Case Blue
            'Blue
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbBlue
                shpRightEdge(i).FillColor = vbBlue
            Next i
            Index = 2
        Case Cyan
            'Cyan
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbCyan
                shpRightEdge(i).FillColor = vbCyan
            Next i
            Index = 3
        Case Green
            'Green
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbGreen
                shpRightEdge(i).FillColor = vbGreen
            Next i
            Index = 4
        Case Magenta
            'Magenta
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbMagenta
                shpRightEdge(i).FillColor = vbMagenta
            Next i
            Index = 5
        Case Red
            'Red
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbRed
                shpRightEdge(i).FillColor = vbRed
            Next i
            Index = 6
        Case Yellow
            'Yellow
            For i = 0 To 19
                shpLeftEdge(i).FillColor = vbYellow
                shpRightEdge(i).FillColor = vbYellow
            Next i
            Index = 7
    End Select

    'Uncheck all Background Color menu items
    For i = 0 To NumberofColors
        mnuEditColorColor(i).Checked = False
    Next i

    'Check menu item that corresponds to background color
    mnuEditColorColor(Index).Checked = True

    Exit Sub

ErrorHandler:

        'If an error occurs, create a brand new file
        'containing the default color of black
        FileNumber = FreeFile

        Open "tcolor.tdf" For Output As #FileNumber
            Write #1, 0
        Close #FileNumber

        'Retry to load color
        LoadColorScheme

End Sub

Private Sub CheckForHighScore()
Dim HighScore As Double
Dim Score As Double
On Error GoTo ErrorHandler

    'Set Score value as listed on status bar
    Score = CDbl(stbScore.Panels(2).Text)

    'Extract current high score values from file
    FileNumber = FreeFile

    Open "ths.tdf" For Input As #FileNumber
        Input #FileNumber, HighName
        Input #FileNumber, HighScore
    Close #FileNumber

    'If the current score is greater than the current high
    'score then call routine that sets the new high score
    If Score > HighScore Then
        SetHighScore Score
    End If

    Exit Sub

ErrorHandler:

    'If error occurs, create a default high score value and
    'try running through the routine again
    CreateDefaultHighScore

    CheckForHighScore

End Sub

Private Sub SetHighScore(Score As Double)
Dim TheName As String
On Error GoTo ErrorHandler

    'Extract current name from file
    FileNumber = FreeFile

    Open "ths.tdf" For Input As #FileNumber
        Input #FileNumber, TheName
    Close #FileNumber

    'Ask user for initials - loop until the number of
    'characters in the entry is less than or equal to 3
    Do
        TheName = InputBox("Congratulations! You have a new high score of: " + CStr(Score) + Chr(13) + "Please enter your name (10 characters max):", "New High Score", TheName)
    Loop Until Len(TheName) <= 10

    'If nothing was entered, set the entry to "N/A"
    If TheName = "" Then
        TheName = "N/A"
    End If

    'Write the new high score data to file
    FileNumber = FreeFile

    Open "ths.tdf" For Output As #FileNumber
        Write #FileNumber, TheName
        Write #FileNumber, Score
    Close #FileNumber

    'Display the new high score data on the status bar
    stbHighScore.Panels(2).Text = CStr(TheName) + " - " + CStr(Score)

    Exit Sub

ErrorHandler:

    'Create a default file and retry routine
    CreateDefaultHighScore

    SetHighScore Score

End Sub

Private Sub SetHighScoreOnBar()
On Error GoTo ErrorHandler

    'Extract high score data from file
    FileNumber = FreeFile

    Open "ths.tdf" For Input As #FileNumber
        Input #FileNumber, TheName
        Input #FileNumber, TheScore
    Close #FileNumber

    'Display the high score on the panel
    stbHighScore.Panels(2).Text = CStr(TheName) + " - " + CStr(TheScore)

    Exit Sub

ErrorHandler:

    'If error occurs, create a default high score value and
    'try running through the routine again
    CreateDefaultHighScore

    SetHighScoreOnBar
End Sub

Private Sub CreateDefaultHighScore()

    'Write default high score of 0 to file
    FileNumber = FreeFile

    Open "ths.tdf" For Output As #FileNumber
        Write #FileNumber, "N/A"
        Write #FileNumber, 0
    Close #FileNumber

End Sub

Private Sub LoadShip()
Dim ShipNumber As Integer
On Error GoTo ErrorHandler

    'Extract ship number from file
    FileNumber = FreeFile

    Open "tship.tdf" For Input As #FileNumber
        Input #FileNumber, ShipNumber
    Close #FileNumber

    'Display the ship that corresponds to the number from
    'the file and uncheck all the menu items
    For i = 0 To NumberofShips
        picShip(i).Visible = False
        mnuEditShipShip(i).Checked = False
    Next i

    picShip(ShipNumber).Visible = True

    'Check the menu item that corresponds to the ship that
    'is displayed
    mnuEditShipShip(ShipNumber).Checked = True

    Exit Sub

ErrorHandler:

    'If error occurs, create file with default ship number
    'and try routine again
    FileNumber = FreeFile

    Open "tship.tdf" For Output As #FileNumber
        Write #FileNumber, 0
    Close #FileNumber

    LoadShip
End Sub

Private Sub LoadSpeed()
Dim TheSpeed As Integer
On Error GoTo ErrorHandler

    'Extract speed from file and apply it to menu and game
    FileNumber = FreeFile

    Open "tspd.tdf" For Input As #FileNumber
        Input #FileNumber, TheSpeed
    Close #FileNumber

    'Change the entry to a string
    TheSpeed = CStr(TheSpeed)

    'Check to make sure that only digits are contained in the entry
    For i = 1 To Len(TheSpeed)
        If Not (Asc(Mid(TheSpeed, i, 1)) >= Asc("0") And Asc(Mid(TheSpeed, i, 1)) <= Asc("9")) Then
            GoTo ErrorHandler
        End If
    Next i

    'Check to make sure the number is between 1 and 1000
    If Not (CDbl(TheSpeed) >= 1 And CDbl(TheSpeed) <= 1000) Then
        GoTo ErrorHandler
    End If

    'Set timer interval
    timMainTimer.Interval = CInt(TheSpeed)

    Exit Sub

ErrorHandler:

    'If error, create default speed and try again
    FileNumber = FreeFile

    Open "tspd.tdf" For Output As #FileNumber
        Write #1, 10
    Close #FileNumber

    LoadSpeed

End Sub
